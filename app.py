# -*- coding: utf-8 -*-
import copy
import json
import os
import re
import threading
import time
import tkinter as tk

import customtkinter as ctk
from tkinter import Menu, filedialog, messagebox, ttk
from PIL import Image, ImageTk

# 分離したモジュールを読み込む
from analysis_engine import (
    DEFAULT_MATERIAL_UNIT_PRICES,
    build_estimation_and_logistics_summary,
    build_export_tables,
    calculate_materials_metrics,
    calculate_volume_and_weight,
    evaluate_formula,
    get_pdf_page_count,
    load_preview_image,
    parse_json_response,
    repair_and_parse_json,
    resolve_material_unit_price,
    run_drawing_analysis,
    safe_float,
    safe_quantity,
)
from api_client import (
    APP_NAME,
    APP_VERSION,
    AnalysisCancelledError,
    DEFAULT_AZURE_API_VERSION,
    DEFAULT_AZURE_DEPLOYMENT_STEP1,
    DEFAULT_AZURE_DEPLOYMENT_STEP2,
    DEFAULT_AZURE_DEPLOYMENT_STEP3,
    DEFAULT_AZURE_MODELS,
    DEFAULT_GEMINI_MODELS,
    DEFAULT_GEMINI_MODEL_STEP1,
    DEFAULT_GEMINI_MODEL_STEP2,
    DEFAULT_GEMINI_MODEL_STEP3,
    DiagnosedAPIError,
    diagnose_api_error,
    get_default_gemini_perf_map,
    parse_azure_endpoint_info,
)
from excel_exporter import open_tables_in_new_excel, save_tables_as_xlsx
from help_dialog import HelpDialog
from settings_dialog import (
    APISettingsDialog,
    AZURE_HEADER_BG,
    BG_COLOR,
    BORDER_COLOR,
    CARD_BG,
    COLOR_DANGER,
    COLOR_DANGER_HOVER,
    COLOR_SUCCESS,
    COLOR_SUCCESS_HOVER,
    MUTED_TEXT,
    PRIMARY,
    PRIMARY_HOVER,
    TEXT_COLOR,
    center_window,
    show_api_error_dialog,
)

# --- 常時ライトモード固定 ---
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# --- 設定ファイルの保存先定義 ---
SETTINGS_DIR = os.path.join(os.path.expanduser("~"), "DrawingAnalysisMiya")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")

# --- 結果タブ名定数 ---
TAB_VOLUME = "📐 体積・重量"
TAB_MATERIALS = "🧱 部材表 (切断寸法)"
TAB_PURCHASED = "🛒 使用購入品"
TAB_MACHINED = "⚙️ 使用機械加工品"
TAB_SUMMARY = "🚛 見積り・搬送サマリー"


# ==============================
# 図面拡大ビューアー（ホイールズーム・ドラッグ移動対応）
# ==============================
class ImageViewerDialog(tk.Toplevel):
    def __init__(self, parent, pil_image, title="図面拡大プレビュー"):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=BG_COLOR)
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        win_w = min(1200, screen_w - 60)
        win_h = min(840, screen_h - 80)
        center_window(self, parent, win_w, win_h)
        self.transient(parent)

        self._orig_img = pil_image.copy()
        self._zoom = 1.0
        self._tk_img = None
        self._img_item = None

        toolbar = tk.Frame(
            self,
            bg=CARD_BG,
            padx=12,
            pady=8,
            highlightbackground=BORDER_COLOR,
            highlightthickness=1,
        )
        toolbar.pack(fill=tk.X, side=tk.TOP)

        ctk.CTkButton(
            toolbar,
            text="➕ 拡大",
            width=80,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            command=lambda: self._apply_zoom(1.25),
        ).pack(side=tk.LEFT, padx=(0, 6))

        ctk.CTkButton(
            toolbar,
            text="➖ 縮小",
            width=80,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            command=lambda: self._apply_zoom(0.8),
        ).pack(side=tk.LEFT, padx=(0, 6))

        ctk.CTkButton(
            toolbar,
            text="🔲 画面に合わせる",
            width=130,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            command=self._fit_to_window,
        ).pack(side=tk.LEFT, padx=(0, 12))

        self._zoom_lbl = tk.Label(
            toolbar,
            text="100%",
            font=("Segoe UI", 9, "bold"),
            bg=CARD_BG,
            fg=PRIMARY,
        )
        self._zoom_lbl.pack(side=tk.LEFT)

        tk.Label(
            toolbar,
            text="操作: マウスホイールで拡大・縮小 ／ ドラッグで移動",
            font=("Meiryo UI", 9),
            bg=CARD_BG,
            fg=MUTED_TEXT,
        ).pack(side=tk.RIGHT)

        canvas_frame = tk.Frame(self, bg=BG_COLOR)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            canvas_frame,
            bg="#E9ECEF",
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            cursor="fleur",
        )
        v_sb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        h_sb = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=v_sb.set, xscrollcommand=h_sb.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        v_sb.grid(row=0, column=1, sticky="ns")
        h_sb.grid(row=1, column=0, sticky="ew")

        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<ButtonPress-1>", lambda e: self.canvas.scan_mark(e.x, e.y))
        self.canvas.bind("<B1-Motion>", lambda e: self.canvas.scan_dragto(e.x, e.y, gain=1))

        self.after(60, self._fit_to_window)

    def _fit_to_window(self):
        cw = max(200, self.canvas.winfo_width() - 20)
        ch = max(200, self.canvas.winfo_height() - 20)
        scale = min(cw / max(1, self._orig_img.width), ch / max(1, self._orig_img.height))
        self._zoom = max(0.1, min(scale, 4.0))
        self._render()

    def _apply_zoom(self, factor):
        self._zoom = max(0.1, min(self._zoom * factor, 6.0))
        self._render()

    def _on_wheel(self, event):
        factor = 1.15 if event.delta > 0 else (1.0 / 1.15)
        self._apply_zoom(factor)
        return "break"

    def _render(self):
        new_w = max(1, int(self._orig_img.width * self._zoom))
        new_h = max(1, int(self._orig_img.height * self._zoom))
        resized = self._orig_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        self._tk_img = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        self._img_item = self.canvas.create_image(0, 0, anchor="nw", image=self._tk_img)
        self.canvas.configure(scrollregion=(0, 0, new_w, new_h))
        self._zoom_lbl.configure(text=f"{int(self._zoom * 100)}%")


# ==============================
# 材質別単価設定ダイアログ（見積り再計算用）
# ==============================
class MaterialPriceDialog(ctk.CTkToplevel):
    def __init__(self, parent, materials_list, current_prices, default_price, on_apply):
        super().__init__(parent)
        self.title("💴 材質別 材料単価設定 (円/kg)")
        self.configure(fg_color=BG_COLOR)
        center_window(self, parent, 480, 520)
        self.transient(parent)
        self.grab_set()

        self._on_apply = on_apply
        self._price_vars = {}
        self._default_var = ctk.StringVar(value=str(int(round(default_price))))

        ctk.CTkLabel(
            self,
            text="材質別 材料単価 (円/kg)",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=PRIMARY,
        ).pack(pady=(12, 4))

        ctk.CTkLabel(
            self,
            text="単価を変更して「適用して再計算」を押すと、AI再通信なしで見積り金額を即更新します。",
            font=ctk.CTkFont(size=11),
            text_color=MUTED_TEXT,
            wraplength=430,
        ).pack(pady=(0, 8))

        def_card = ctk.CTkFrame(
            self, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR
        )
        def_card.pack(fill="x", padx=16, pady=(0, 8))
        ctk.CTkLabel(
            def_card,
            text="基本単価 (未登録の材質に適用):",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=12, pady=8)
        ctk.CTkEntry(
            def_card,
            textvariable=self._default_var,
            width=90,
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        ).pack(side="right", padx=12, pady=8)

        scroll = ctk.CTkScrollableFrame(
            self, fg_color=CARD_BG, border_width=1, border_color=BORDER_COLOR
        )
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        # 抽出された材質を先頭に、代表的な材質を続けて表示
        ordered_mats = []
        for m in materials_list:
            if m and m not in ordered_mats:
                ordered_mats.append(m)
        for m in ("SS400", "S45C", "SUS304", "SUS316", "A5052", "FC250", "MCナイロン"):
            if m not in ordered_mats:
                ordered_mats.append(m)

        for mat in ordered_mats:
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", padx=6, pady=3)
            val = resolve_material_unit_price(mat, current_prices, default_price)
            var = ctk.StringVar(value=str(int(round(val))))
            self._price_vars[mat] = var
            ctk.CTkLabel(
                row,
                text=mat,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=TEXT_COLOR,
                anchor="w",
            ).pack(side="left", padx=(4, 8))
            ctk.CTkLabel(row, text="円/kg", font=ctk.CTkFont(size=11), text_color=MUTED_TEXT).pack(
                side="right", padx=(4, 4)
            )
            ctk.CTkEntry(
                row,
                textvariable=var,
                width=95,
                fg_color="white",
                border_color="#CED4DA",
                text_color=TEXT_COLOR,
            ).pack(side="right")

        btn_bar = ctk.CTkFrame(self, fg_color="transparent")
        btn_bar.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkButton(
            btn_bar,
            text="初期目安に戻す",
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            width=120,
            command=self._reset_defaults,
        ).pack(side="left")

        ctk.CTkButton(
            btn_bar,
            text="適用して再計算",
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            font=ctk.CTkFont(weight="bold"),
            width=150,
            command=self._apply,
        ).pack(side="right")

        ctk.CTkButton(
            btn_bar,
            text="キャンセル",
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            width=95,
            command=self.destroy,
        ).pack(side="right", padx=(0, 8))

    def _reset_defaults(self):
        self._default_var.set("250")
        for mat, var in self._price_vars.items():
            val = resolve_material_unit_price(mat, {}, 250.0)
            var.set(str(int(round(val))))

    def _apply(self):
        new_default = max(0.0, safe_float(self._default_var.get()) or 250.0)
        new_prices = {}
        for mat, var in self._price_vars.items():
            new_prices[mat] = max(0.0, safe_float(var.get()))
        self._on_apply(new_prices, new_default)
        self.destroy()


# --- メインアプリケーションクラス ---
class DrawingAnalysisApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(f"{APP_NAME} {APP_VERSION} - 2D機械図面 解析ツール")
        self.geometry("1200x800")
        self.minsize(960, 680)
        self.configure(fg_color=BG_COLOR)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.selected_file_path = None
        self.selected_page_index = 0
        self.pdf_page_count = 1
        self.processed_image_pil = None
        self.last_result_data = None
        self._cancel_event = None
        self._is_analyzing = False
        self._analysis_start_time = None
        self._current_progress_ratio = 0.0
        self._elapsed_timer_id = None

        self.default_free_perf = get_default_gemini_perf_map(DEFAULT_GEMINI_MODELS, is_free=True)
        self.default_paid_perf = get_default_gemini_perf_map(DEFAULT_GEMINI_MODELS, is_free=False)

        self.api_settings = {
            "engine": "gemini",
            "plan": "free",
            "free_key": "",
            "paid_key": "",
            "free_model_step1": DEFAULT_GEMINI_MODEL_STEP1,
            "free_model_step2": DEFAULT_GEMINI_MODEL_STEP2,
            "free_model_step3": DEFAULT_GEMINI_MODEL_STEP3,
            "paid_model_step1": DEFAULT_GEMINI_MODEL_STEP1,
            "paid_model_step2": DEFAULT_GEMINI_MODEL_STEP2,
            "paid_model_step3": DEFAULT_GEMINI_MODEL_STEP3,
            "free_model_perf": copy.deepcopy(self.default_free_perf),
            "paid_model_perf": copy.deepcopy(self.default_paid_perf),
            "free_temp": 0.0,
            "paid_temp": 0.0,
            "free_safety": True,
            "paid_safety": True,
            "free_tokens": 65535,
            "paid_tokens": 65535,
            "free_prompts": [],
            "paid_prompts": [],
            "models_list": [list(m) for m in DEFAULT_GEMINI_MODELS],
            "gemini_model_meta": {},
            "gemini_mainstream_only": True,
            # Azure OpenAI Service 設定
            "azure_endpoint": "",
            "azure_api_key": "",
            "azure_deployment_step1": DEFAULT_AZURE_DEPLOYMENT_STEP1,
            "azure_deployment_step2": DEFAULT_AZURE_DEPLOYMENT_STEP2,
            "azure_deployment_step3": DEFAULT_AZURE_DEPLOYMENT_STEP3,
            "azure_api_version": DEFAULT_AZURE_API_VERSION,
            "azure_rpm": 60,
            "azure_threads": 3,
            "azure_temp": 0.0,
            "azure_tokens": 16384,
            "azure_prompts": [],
            "azure_models_list": list(DEFAULT_AZURE_MODELS),
            "azure_model_meta": {},
            "azure_mainstream_only": True,
            # 共通お気に入り
            "saved_prompts": [],
            # 解析項目の選択状態
            "analysis_targets": {
                "volume": True,
                "materials": True,
                "purchased_parts": True,
                "machined_parts": True,
            },
            # 基本比重・見積り用材料単価設定
            "default_density": 7.85,
            "default_unit_price": 250.0,
            "material_unit_prices": copy.deepcopy(DEFAULT_MATERIAL_UNIT_PRICES),
        }

        self.load_settings()

        saved_targets = self.api_settings.get("analysis_targets", {}) or {}
        self.target_vars = {
            "volume": ctk.BooleanVar(value=bool(saved_targets.get("volume", True))),
            "materials": ctk.BooleanVar(value=bool(saved_targets.get("materials", True))),
            "purchased_parts": ctk.BooleanVar(
                value=bool(saved_targets.get("purchased_parts", True))
            ),
            "machined_parts": ctk.BooleanVar(
                value=bool(saved_targets.get("machined_parts", True))
            ),
        }

        self._sort_states = {}
        self._setup_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(200, self._maximize_window)

    def _on_close(self):
        """ウィンドウ終了時に実行中の解析・タイマーを安全に停止してから終了する"""
        if self._cancel_event is not None:
            self._cancel_event.set()
        self._stop_elapsed_timer()
        self.destroy()

    def _get_effective_default_unit_price(self):
        """入力欄の基本単価を取得する（0円も有効値として扱い、空欄・不正値のみ設定値へフォールバック）"""
        if hasattr(self, "unit_price_entry"):
            raw = str(self.unit_price_entry.get()).strip().replace(",", "")
            if raw:
                try:
                    val = float(raw)
                    if val >= 0.0:
                        return val
                except ValueError:
                    pass
        return max(0.0, safe_float(self.api_settings.get("default_unit_price", 250.0)))

    def _set_page_selector(self, values, selected_label, enabled):
        """disabled 状態での set() 無視バグを防ぐため、必ず normal で値を設定してから状態を反映する"""
        self.page_selector.configure(state="normal", values=values)
        self.page_selector.set(selected_label)
        self.page_selector.configure(state="readonly" if enabled else "disabled")

    def _safe_after(self, delay_ms, fn):
        try:
            if self.winfo_exists():
                self.after(delay_ms, lambda: self.winfo_exists() and fn())
        except Exception:
            pass

    def _maximize_window(self):
        self.update_idletasks()
        try:
            self.state("zoomed")
        except Exception:
            pass

    def load_settings(self):
        """設定を読み込み、過去の設定形式からの自動移行も行う"""
        if not os.path.exists(SETTINGS_FILE):
            return
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)

            if not isinstance(loaded, dict):
                return

            for plan in ("free", "paid"):
                old_model_key = f"{plan}_model"
                if old_model_key in loaded:
                    old_val = str(loaded[old_model_key])
                    if "1.5" in old_val or "2.0" in old_val:
                        old_val = DEFAULT_GEMINI_MODEL_STEP1
                    for step in ("step1", "step2", "step3"):
                        step_key = f"{plan}_model_{step}"
                        if step_key not in loaded:
                            loaded[step_key] = old_val
                    del loaded[old_model_key]

            self.api_settings.update(loaded)

            if self.api_settings.get("engine") not in ("gemini", "azure"):
                self.api_settings["engine"] = "gemini"
            if self.api_settings.get("plan") not in ("free", "paid"):
                self.api_settings["plan"] = "free"

            for plan, def_perf in (
                ("free", self.default_free_perf),
                ("paid", self.default_paid_perf),
            ):
                perf_key = f"{plan}_model_perf"
                if not isinstance(self.api_settings.get(perf_key), dict):
                    self.api_settings[perf_key] = copy.deepcopy(def_perf)
                else:
                    for m_id, vals in def_perf.items():
                        self.api_settings[perf_key].setdefault(m_id, copy.deepcopy(vals))

            protected_keys = {"free_model_perf", "paid_model_perf", "azure_rpm", "azure_threads"}
            keys_to_delete = [
                k
                for k in list(self.api_settings.keys())
                if ("rpm" in k or "threads" in k) and k not in protected_keys
            ]
            for k in keys_to_delete:
                del self.api_settings[k]

            for plan in ("free", "paid"):
                for step, def_m in (
                    ("step1", DEFAULT_GEMINI_MODEL_STEP1),
                    ("step2", DEFAULT_GEMINI_MODEL_STEP2),
                    ("step3", DEFAULT_GEMINI_MODEL_STEP3),
                ):
                    step_key = f"{plan}_model_{step}"
                    cur_val = str(self.api_settings.get(step_key, ""))
                    if not cur_val or "1.5" in cur_val or "2.0" in cur_val:
                        self.api_settings[step_key] = def_m

            saved_list = self.api_settings.get("models_list", [])
            if not saved_list or any(
                not isinstance(item, (list, tuple)) or len(item) < 2 or "1.5" in str(item[1])
                for item in saved_list
            ):
                self.api_settings["models_list"] = [list(m) for m in DEFAULT_GEMINI_MODELS]

            if not self.api_settings.get("azure_deployment_step1"):
                self.api_settings["azure_deployment_step1"] = DEFAULT_AZURE_DEPLOYMENT_STEP1
            if not self.api_settings.get("azure_deployment_step2"):
                self.api_settings["azure_deployment_step2"] = self.api_settings[
                    "azure_deployment_step1"
                ]
            if not self.api_settings.get("azure_deployment_step3"):
                self.api_settings["azure_deployment_step3"] = DEFAULT_AZURE_DEPLOYMENT_STEP3

            if not isinstance(self.api_settings.get("analysis_targets"), dict):
                self.api_settings["analysis_targets"] = {
                    "volume": True,
                    "materials": True,
                    "purchased_parts": True,
                    "machined_parts": True,
                }
            else:
                for k in ("volume", "materials", "purchased_parts", "machined_parts"):
                    self.api_settings["analysis_targets"].setdefault(k, True)

            if not isinstance(self.api_settings.get("material_unit_prices"), dict):
                self.api_settings["material_unit_prices"] = copy.deepcopy(
                    DEFAULT_MATERIAL_UNIT_PRICES
                )
            else:
                for k, v in DEFAULT_MATERIAL_UNIT_PRICES.items():
                    self.api_settings["material_unit_prices"].setdefault(k, v)

            den_loaded = safe_float(self.api_settings.get("default_density", 7.85))
            self.api_settings["default_density"] = den_loaded if den_loaded > 0.0 else 7.85
            price_loaded = safe_float(self.api_settings.get("default_unit_price", 250.0))
            self.api_settings["default_unit_price"] = max(0.0, price_loaded)

        except Exception as e:
            print(f"Load settings error: {e}")

    def save_settings(self):
        """設定を保存する"""
        try:
            if hasattr(self, "target_vars"):
                self.api_settings["analysis_targets"] = self._get_selected_targets()
            if hasattr(self, "density_entry"):
                den_val = safe_float(self.density_entry.get())
                if den_val > 0.0:
                    self.api_settings["default_density"] = den_val
            if hasattr(self, "unit_price_entry"):
                raw_price = str(self.unit_price_entry.get()).strip()
                if raw_price:
                    try:
                        p_val = float(raw_price.replace(",", ""))
                        if p_val >= 0.0:
                            self.api_settings["default_unit_price"] = p_val
                    except ValueError:
                        pass
            os.makedirs(SETTINGS_DIR, exist_ok=True)
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.api_settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Save settings error: {e}")

    @staticmethod
    def _format_density(val):
        d = safe_float(val)
        if d <= 0.0:
            d = 7.85
        return f"{d:g}"

    def save_basic_params(self, explicit=True):
        """基本比重・基本単価を検証して設定ファイルへ保存し、表示中の見積りへも即反映する"""
        raw_den = str(self.density_entry.get()).strip().replace(",", "")
        raw_price = str(self.unit_price_entry.get()).strip().replace(",", "")
        try:
            den_val = float(raw_den)
            if den_val <= 0.0:
                raise ValueError("density must be positive")
        except ValueError:
            if explicit:
                messagebox.showwarning(
                    "入力エラー",
                    "基本比重には 0 より大きい数値（例: 7.85）を入力してください。",
                    parent=self,
                )
            return False

        try:
            price_val = float(raw_price)
            if price_val < 0.0:
                raise ValueError("price must be non-negative")
        except ValueError:
            if explicit:
                messagebox.showwarning(
                    "入力エラー",
                    "基本単価には 0 以上の数値（例: 250）を入力してください。",
                    parent=self,
                )
            return False

        prev_den = safe_float(self.api_settings.get("default_density", 7.85))
        prev_price = safe_float(self.api_settings.get("default_unit_price", 250.0))
        changed = (abs(den_val - prev_den) > 1e-6) or (abs(price_val - prev_price) > 1e-6)

        self.api_settings["default_density"] = den_val
        self.api_settings["default_unit_price"] = price_val

        if changed or explicit:
            self.save_settings()
            if self.last_result_data:
                self.last_result_data["default_density"] = den_val
                self._refresh_summary_views(self.last_result_data)
            self.update_status(
                f"基本比重 ({self._format_density(den_val)} g/cm³)・基本単価 ({int(round(price_val)):,} 円/kg) を保存しました"
            )
            if explicit and hasattr(self, "save_params_btn"):
                self.save_params_btn.configure(text="✅ 保存しました", fg_color="#D1E7DD", text_color="#0F5132")
                self._safe_after(
                    1600,
                    lambda: self.save_params_btn.configure(
                        text="💾 比重・単価を保存", fg_color="#E7F1FF", text_color=PRIMARY
                    ),
                )
        return True

    def _get_selected_targets(self):
        return {k: bool(v.get()) for k, v in self.target_vars.items()}

    def _on_target_toggle(self):
        self.save_settings()

    def _profile_label_from_settings(self):
        if self.api_settings.get("engine") == "azure":
            return "Azure OpenAI"
        return "Gemini 課金" if self.api_settings.get("plan") == "paid" else "Gemini 無料"

    def _on_quick_profile_change(self, selected_label):
        if selected_label == "Azure OpenAI":
            self.api_settings["engine"] = "azure"
        elif selected_label == "Gemini 課金":
            self.api_settings["engine"] = "gemini"
            self.api_settings["plan"] = "paid"
        else:
            self.api_settings["engine"] = "gemini"
            self.api_settings["plan"] = "free"
        self._refresh_api_panel()
        self.save_settings()

    def _refresh_api_panel(self):
        engine = self.api_settings.get("engine", "gemini")
        self.profile_selector.set(self._profile_label_from_settings())

        if engine == "azure":
            dep1 = self.api_settings.get("azure_deployment_step1", "").strip()
            dep3 = self.api_settings.get("azure_deployment_step3", "").strip() or dep1
            dep_disp = dep1 if dep1 == dep3 else f"{dep1} / {dep3}"
            self.plan_indicator.configure(
                text=f"🔷 Azure OpenAI Service\nデプロイ: {dep_disp or '未設定'}",
                text_color=PRIMARY,
            )
            has_key = bool(self.api_settings.get("azure_api_key", "").strip())
            has_ep = bool(self.api_settings.get("azure_endpoint", "").strip())
            if has_key and has_ep and dep1:
                self.cred_status_label.configure(
                    text="接続情報: 設定済み ✅",
                    text_color=COLOR_SUCCESS,
                )
            else:
                missing = []
                if not has_ep:
                    missing.append("URL")
                if not has_key:
                    missing.append("キー")
                if not dep1:
                    missing.append("デプロイ名")
                self.cred_status_label.configure(
                    text=f"接続情報: 未設定 ({'・'.join(missing)})",
                    text_color=COLOR_DANGER,
                )
        else:
            plan = self.api_settings.get("plan", "free")
            m1 = self.api_settings.get(f"{plan}_model_step1", DEFAULT_GEMINI_MODEL_STEP1)
            m3 = self.api_settings.get(f"{plan}_model_step3", DEFAULT_GEMINI_MODEL_STEP3)
            model_info = m1 if m1 == m3 else f"{m1} / {m3}"
            if plan == "free":
                self.plan_indicator.configure(
                    text=f"🟢 Gemini API (無料枠)\nモデル: {model_info}",
                    text_color=COLOR_SUCCESS,
                )
            else:
                self.plan_indicator.configure(
                    text=f"🔵 Gemini API (課金枠)\nモデル: {model_info}",
                    text_color=PRIMARY,
                )
            has_key = bool(self.api_settings.get(f"{plan}_key", "").strip())
            if has_key:
                self.cred_status_label.configure(
                    text="APIキー: 設定済み ✅",
                    text_color=COLOR_SUCCESS,
                )
            else:
                self.cred_status_label.configure(
                    text="APIキー: 未設定（詳細設定から入力）",
                    text_color=COLOR_DANGER,
                )

    def _sort_treeview_column(self, tree, col_id):
        key = (id(tree), col_id)
        reverse = self._sort_states.get(key, False)

        def _nat_key(val):
            s = str(val or "").strip().replace(",", "")
            return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]

        rows = [(tree.set(item_id, col_id), item_id) for item_id in tree.get_children("")]
        rows.sort(key=lambda pair: _nat_key(pair[0]), reverse=reverse)
        for index, (_, item_id) in enumerate(rows):
            tree.move(item_id, "", index)
        self._sort_states[key] = not reverse

    def _bind_tree_copy_menu(self, tree, headers):
        def _collect_all_rows(parent=""):
            out = []
            for iid in tree.get_children(parent):
                out.append([str(v) for v in tree.item(iid, "values")])
                out.extend(_collect_all_rows(iid))
            return out

        def _copy_selected():
            sels = tree.selection()
            if not sels:
                return
            lines = ["\t".join(str(v) for v in tree.item(iid, "values")) for iid in sels]
            try:
                self.clipboard_clear()
                self.clipboard_append("\n".join(lines))
                self.update_status("選択行をクリップボードにコピーしました")
            except Exception:
                pass

        def _copy_all():
            rows = _collect_all_rows("")
            if not rows:
                return
            lines = ["\t".join(headers)] + ["\t".join(r) for r in rows]
            try:
                self.clipboard_clear()
                self.clipboard_append("\n".join(lines))
                self.update_status("表全体をクリップボードにコピーしました（Excel等へ貼り付け可能）")
            except Exception:
                pass

        def _show_menu(event):
            row_id = tree.identify_row(event.y)
            if row_id and row_id not in tree.selection():
                tree.selection_set(row_id)
            menu = Menu(tree, tearoff=0, bg="white", fg=TEXT_COLOR)
            menu.add_command(label="選択行をコピー", command=_copy_selected)
            menu.add_command(label="表全体をコピー (Excel貼り付け用)", command=_copy_all)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        tree.bind("<Button-3>", _show_menu)

    def _create_result_treeview(self, parent_tab, columns_spec, sortable=False, height=None):
        frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=4, pady=4)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        col_ids = tuple(c[0] for c in columns_spec)
        headers = [c[1] for c in columns_spec]
        tv_kwargs = {"columns": col_ids, "show": "headings"}
        if height is not None:
            tv_kwargs["height"] = height
        tree = ttk.Treeview(frame, **tv_kwargs)

        for col_id, text, width, anchor in columns_spec:
            if sortable:
                tree.heading(
                    col_id,
                    text=text,
                    command=lambda t=tree, c=col_id: self._sort_treeview_column(t, c),
                )
            else:
                tree.heading(col_id, text=text)
            tree.column(col_id, width=width, minwidth=45, anchor=anchor)

        v_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        self._bind_tree_copy_menu(tree, headers)
        return tree

    def _build_summary_tab(self, tab_summary):
        """見積り（材質別重量・金額）・最大外形寸法・搬送（推奨トラック）・定尺発注集計ダッシュボードを構築"""
        scroll = ctk.CTkScrollableFrame(tab_summary, fg_color=BG_COLOR)
        scroll.pack(fill="both", expand=True, padx=2, pady=2)

        # 0. 製品最大外形寸法（最大長さ・最大幅・最大高さ）＆ 総重量 KPI カード
        dim_card = ctk.CTkFrame(
            scroll, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR
        )
        dim_card.pack(fill="x", padx=6, pady=(4, 8))

        ctk.CTkLabel(
            dim_card,
            text="📏 図面の製品最大外形寸法 ＆ 総重量",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        kpi_grid = ctk.CTkFrame(dim_card, fg_color="transparent")
        kpi_grid.pack(fill="x", padx=10, pady=(0, 10))
        for col_i in range(4):
            kpi_grid.grid_columnconfigure(col_i, weight=1)

        def _make_kpi_box(col_i, caption, init_val, val_color=PRIMARY):
            box = ctk.CTkFrame(
                kpi_grid,
                fg_color="#F8FAFC",
                corner_radius=6,
                border_width=1,
                border_color=BORDER_COLOR,
            )
            box.grid(row=0, column=col_i, sticky="ew", padx=4, pady=2)
            ctk.CTkLabel(
                box,
                text=caption,
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color=MUTED_TEXT,
            ).pack(anchor="w", padx=10, pady=(6, 0))
            val_lbl = ctk.CTkLabel(
                box,
                text=init_val,
                font=ctk.CTkFont(size=16, weight="bold"),
                text_color=val_color,
            )
            val_lbl.pack(anchor="w", padx=10, pady=(2, 6))
            return val_lbl

        self.dim_length_label = _make_kpi_box(0, "最大長さ (L)", "--- mm")
        self.dim_width_label = _make_kpi_box(1, "最大幅 (W)", "--- mm")
        self.dim_height_label = _make_kpi_box(2, "最大高さ (H)", "--- mm")
        self.dim_weight_label = _make_kpi_box(3, "総重量", "--- kg", val_color=COLOR_SUCCESS)

        # 1. 搬送・トラック選定カード
        truck_card = ctk.CTkFrame(
            scroll, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR
        )
        truck_card.pack(fill="x", padx=6, pady=(0, 8))

        truck_head = ctk.CTkFrame(truck_card, fg_color="transparent")
        truck_head.pack(fill="x", padx=12, pady=(8, 4))
        ctk.CTkLabel(
            truck_head,
            text="🚛 搬送・推奨トラック選定目安",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=PRIMARY,
        ).pack(side="left")

        self.truck_badge = ctk.CTkLabel(
            truck_head,
            text=" 解析後に自動判定 ",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=AZURE_HEADER_BG,
            text_color=PRIMARY,
            corner_radius=6,
            padx=10,
            pady=3,
        )
        self.truck_badge.pack(side="right")

        self.truck_detail_label = ctk.CTkLabel(
            truck_card,
            text="総重量・製品全体外形寸法 (L×W×H) ・最大長尺部材から推奨トラックと荷役目安を算出します。",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_COLOR,
            justify="left",
            anchor="w",
            wraplength=760,
        )
        self.truck_detail_label.pack(fill="x", padx=12, pady=(2, 10))

        # 2. 材質別 重量集計 ＆ 概算材料費カード
        cost_card = ctk.CTkFrame(
            scroll, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR
        )
        cost_card.pack(fill="x", padx=6, pady=(0, 8))

        cost_head = ctk.CTkFrame(cost_card, fg_color="transparent")
        cost_head.pack(fill="x", padx=12, pady=(8, 4))
        ctk.CTkLabel(
            cost_head,
            text="💰 材質別 重量集計 ＆ 概算材料費",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=PRIMARY,
        ).pack(side="left")

        self.price_edit_btn = ctk.CTkButton(
            cost_head,
            text="💴 単価設定・再計算",
            width=140,
            height=26,
            fg_color="#E7F1FF",
            hover_color="#CFE2FF",
            text_color=PRIMARY,
            border_width=1,
            border_color="#B6D4FE",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.open_material_price_dialog,
        )
        self.price_edit_btn.pack(side="right")

        self.cost_total_label = ctk.CTkLabel(
            cost_head,
            text="概算材料費合計: --- 円",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLOR_SUCCESS,
        )
        self.cost_total_label.pack(side="right", padx=(0, 14))

        cost_table_box = ctk.CTkFrame(cost_card, fg_color="transparent")
        cost_table_box.pack(fill="x", padx=8, pady=(0, 8))
        self.cost_summary_tree = self._create_result_treeview(
            cost_table_box,
            [
                ("material", "材質", 130, "w"),
                ("density", "比重 (g/cm³)", 95, "e"),
                ("weight", "合計重量 (kg)", 120, "e"),
                ("unit_price", "適用単価 (円/kg)", 120, "e"),
                ("cost", "概算材料費 (円)", 140, "e"),
            ],
            sortable=True,
            height=5,
        )

        # 3. 鋼材・定尺発注集計カード
        stock_card = ctk.CTkFrame(
            scroll, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR
        )
        stock_card.pack(fill="x", padx=6, pady=(0, 8))

        ctk.CTkLabel(
            stock_card,
            text="📏 鋼材・断面規格別 定尺発注集計（部材表の名寄せ・ロス10%見込み）",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        stock_table_box = ctk.CTkFrame(stock_card, fg_color="transparent")
        stock_table_box.pack(fill="x", padx=8, pady=(0, 8))
        self.stock_summary_tree = self._create_result_treeview(
            stock_table_box,
            [
                ("material_type", "部材種別", 110, "w"),
                ("material_spec", "材質", 90, "w"),
                ("section_size", "断面規格・板厚", 130, "w"),
                ("pieces", "切断本数(枚)", 85, "e"),
                ("total_len", "合計切断長 (mm)", 115, "e"),
                ("stock_guide", "定尺発注目安 (5.5m / 6.0m)", 215, "w"),
                ("weight", "小計重量 (kg)", 95, "e"),
            ],
            sortable=True,
            height=5,
        )

    def _setup_ui(self):
        # メニューバー（ヘルプメニュー）
        menubar = Menu(self, bg="white", fg=TEXT_COLOR)
        help_menu = Menu(menubar, tearoff=0, bg="white", fg=TEXT_COLOR)
        help_menu.add_command(
            label="🚀 使い方ガイド (基本操作)", command=lambda: self.open_help("guide")
        )
        help_menu.add_command(
            label="🤖 AI・API の準備 (Gemini / Azure)", command=lambda: self.open_help("ai")
        )
        help_menu.add_command(
            label="📐 解析精度を上げるコツ", command=lambda: self.open_help("tips")
        )
        help_menu.add_command(
            label="❓ よくある質問・トラブル解決 (FAQ)", command=lambda: self.open_help("faq")
        )
        help_menu.add_separator()
        help_menu.add_command(
            label="📄 Readme / アプリ情報", command=lambda: self.open_help("about")
        )
        menubar.add_cascade(label="ヘルプ", menu=help_menu)
        self.config(menu=menubar)

        # 左側パネル
        self.left_frame = ctk.CTkFrame(
            self,
            width=335,
            corner_radius=10,
            fg_color=CARD_BG,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        self.left_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.left_frame.grid_rowconfigure(0, weight=1)
        self.left_frame.grid_columnconfigure(0, weight=1)

        self.left_scroll = ctk.CTkScrollableFrame(
            self.left_frame,
            width=305,
            fg_color="transparent",
            corner_radius=0,
        )
        self.left_scroll.grid(row=0, column=0, sticky="nsew", padx=4, pady=(6, 0))

        # AI / API 設定エリア ＋ ヘルプボタン
        top_header_row = ctk.CTkFrame(self.left_scroll, fg_color="transparent")
        top_header_row.pack(fill="x", padx=14, pady=(8, 4))

        ctk.CTkLabel(
            top_header_row,
            text="AI / API 設定:",
            font=ctk.CTkFont(weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left")

        self.help_btn = ctk.CTkButton(
            top_header_row,
            text="❓ ヘルプ",
            width=78,
            height=26,
            fg_color="#E7F1FF",
            hover_color="#CFE2FF",
            text_color=PRIMARY,
            border_width=1,
            border_color="#B6D4FE",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self.open_help("guide"),
        )
        self.help_btn.pack(side="right")

        self.profile_selector = ctk.CTkSegmentedButton(
            self.left_scroll,
            values=["Gemini 無料", "Gemini 課金", "Azure OpenAI"],
            command=self._on_quick_profile_change,
            fg_color="#DEE2E6",
            selected_color=PRIMARY,
            selected_hover_color=PRIMARY_HOVER,
            unselected_color="#6C757D",
            unselected_hover_color="#5C636A",
            text_color="white",
            font=ctk.CTkFont(size=11, weight="bold"),
        )
        self.profile_selector.pack(fill="x", padx=14, pady=(0, 6))

        self.api_settings_btn = ctk.CTkButton(
            self.left_scroll,
            text="⚙️ AI詳細設定 (APIキー / モデル)",
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            command=self.open_api_settings,
        )
        self.api_settings_btn.pack(fill="x", padx=14, pady=(0, 6))

        self.api_status_card = ctk.CTkFrame(
            self.left_scroll,
            fg_color="#F8FAFC",
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        self.api_status_card.pack(fill="x", padx=14, pady=(0, 10))

        self.plan_indicator = ctk.CTkLabel(
            self.api_status_card,
            text="",
            font=ctk.CTkFont(size=12, weight="bold"),
            justify="left",
            wraplength=250,
        )
        self.plan_indicator.pack(anchor="w", padx=10, pady=(6, 2))

        self.cred_status_label = ctk.CTkLabel(
            self.api_status_card,
            text="",
            font=ctk.CTkFont(size=11, weight="bold"),
            justify="left",
            wraplength=250,
        )
        self.cred_status_label.pack(anchor="w", padx=10, pady=(0, 6))

        self._refresh_api_panel()

        # ファイル選択
        ctk.CTkLabel(
            self.left_scroll,
            text="図面ファイル:",
            font=ctk.CTkFont(weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=14, pady=(4, 4))
        self.select_btn = ctk.CTkButton(
            self.left_scroll,
            text="ファイルを選択",
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            command=self.select_file,
        )
        self.select_btn.pack(fill="x", padx=14, pady=(0, 4))
        self.file_path_display = ctk.CTkLabel(
            self.left_scroll, text="未選択", text_color=MUTED_TEXT, wraplength=255
        )
        self.file_path_display.pack(anchor="w", padx=14, pady=(0, 6))

        # 複数ページ PDF 用ページセレクタ
        self.page_row = ctk.CTkFrame(self.left_scroll, fg_color="transparent")
        self.page_row.pack(fill="x", padx=14, pady=(0, 8))
        ctk.CTkLabel(
            self.page_row,
            text="対象ページ:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=(0, 8))
        self.page_selector = ctk.CTkComboBox(
            self.page_row,
            values=["1 / 1"],
            width=120,
            fg_color="white",
            border_color="#CED4DA",
            button_color="#E9ECEF",
            button_hover_color="#DEE2E6",
            dropdown_fg_color="white",
            dropdown_hover_color="#E7F1FF",
            dropdown_text_color=TEXT_COLOR,
            text_color=TEXT_COLOR,
            command=self._on_page_selected,
        )
        self._set_page_selector(["1 / 1"], "1 / 1", enabled=False)
        self.page_selector.pack(side="left", fill="x", expand=True)

        # 解析項目の選択（複数選択可）
        ctk.CTkLabel(
            self.left_scroll,
            text="解析項目の選択:",
            font=ctk.CTkFont(weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=14, pady=(4, 4))

        targets_card = ctk.CTkFrame(
            self.left_scroll,
            fg_color="#F8FAFC",
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        targets_card.pack(fill="x", padx=14, pady=(0, 10))
        targets_card.grid_columnconfigure(0, weight=1)
        targets_card.grid_columnconfigure(1, weight=1)

        target_defs = [
            ("volume", "体積・重量・搬送", 0, 0),
            ("materials", "部材表・切断寸法", 0, 1),
            ("purchased_parts", "使用購入品", 1, 0),
            ("machined_parts", "使用機械加工品", 1, 1),
        ]
        for key, label, r_idx, c_idx in target_defs:
            ctk.CTkCheckBox(
                targets_card,
                text=label,
                variable=self.target_vars[key],
                command=self._on_target_toggle,
                text_color=TEXT_COLOR,
                fg_color=PRIMARY,
                hover_color=PRIMARY_HOVER,
                border_color="#ADB5BD",
                checkbox_width=16,
                checkbox_height=16,
                font=ctk.CTkFont(size=11, weight="bold"),
            ).grid(row=r_idx, column=c_idx, sticky="w", padx=7, pady=5)

        # 基本比重 & 材料単価目安
        param_grid = ctk.CTkFrame(self.left_scroll, fg_color="transparent")
        param_grid.pack(fill="x", padx=14, pady=(2, 8))
        param_grid.grid_columnconfigure(0, weight=1)
        param_grid.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            param_grid,
            text="基本比重 (g/cm³):",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=TEXT_COLOR,
        ).grid(row=0, column=0, sticky="w", padx=(0, 4))
        ctk.CTkLabel(
            param_grid,
            text="基本単価 (円/kg):",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=TEXT_COLOR,
        ).grid(row=0, column=1, sticky="w", padx=(4, 0))

        self.density_entry = ctk.CTkEntry(
            param_grid,
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        )
        self.density_entry.insert(
            0, self._format_density(self.api_settings.get("default_density", 7.85))
        )
        self.density_entry.grid(row=1, column=0, sticky="ew", padx=(0, 4), pady=(2, 0))
        self.density_entry.bind("<FocusOut>", lambda e: self.save_basic_params(explicit=False))
        self.density_entry.bind("<Return>", lambda e: self.save_basic_params(explicit=True))

        self.unit_price_entry = ctk.CTkEntry(
            param_grid,
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        )
        self.unit_price_entry.insert(
            0, str(int(round(safe_float(self.api_settings.get("default_unit_price", 250.0)))))
        )
        self.unit_price_entry.grid(row=1, column=1, sticky="ew", padx=(4, 0), pady=(2, 0))
        self.unit_price_entry.bind("<FocusOut>", lambda e: self.save_basic_params(explicit=False))
        self.unit_price_entry.bind("<Return>", lambda e: self.save_basic_params(explicit=True))

        self.save_params_btn = ctk.CTkButton(
            param_grid,
            text="💾 比重・単価を保存",
            height=26,
            fg_color="#E7F1FF",
            hover_color="#CFE2FF",
            text_color=PRIMARY,
            border_width=1,
            border_color="#B6D4FE",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self.save_basic_params(explicit=True),
        )
        self.save_params_btn.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        # 下部固定アクションエリア
        self.left_bottom = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.left_bottom.grid(row=1, column=0, sticky="ew", padx=18, pady=(6, 14))

        self.analyze_btn = ctk.CTkButton(
            self.left_bottom,
            text="解析を実行",
            fg_color=COLOR_SUCCESS,
            hover_color=COLOR_SUCCESS_HOVER,
            text_color="white",
            font=ctk.CTkFont(weight="bold"),
            command=self.on_analyze_or_cancel_click,
        )
        self.analyze_btn.pack(fill="x", pady=(0, 8))

        self.export_btn = ctk.CTkButton(
            self.left_bottom,
            text="📊 Excelで開く",
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            font=ctk.CTkFont(weight="bold"),
            command=self.open_in_excel,
            state="disabled",
        )
        self.export_btn.pack(fill="x", pady=(0, 8))

        # 実行進捗カード（フェーズ・進捗率・経過時間・プログレスバー・詳細ステータス）
        self.progress_card = ctk.CTkFrame(
            self.left_bottom,
            fg_color="#F8FAFC",
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        self.progress_card.pack(fill="x", pady=(2, 0))

        prog_top = ctk.CTkFrame(self.progress_card, fg_color="transparent")
        prog_top.pack(fill="x", padx=10, pady=(7, 3))

        self.progress_phase_label = ctk.CTkLabel(
            prog_top,
            text="待機中",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=PRIMARY,
            anchor="w",
        )
        self.progress_phase_label.pack(side="left", fill="x", expand=True)

        self.progress_meta_label = ctk.CTkLabel(
            prog_top,
            text="0%",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=MUTED_TEXT,
            anchor="e",
        )
        self.progress_meta_label.pack(side="right", padx=(6, 0))

        self.progress_bar = ctk.CTkProgressBar(
            self.progress_card,
            height=10,
            corner_radius=5,
            fg_color="#DEE2E6",
            progress_color=PRIMARY,
        )
        self.progress_bar.set(0.0)
        self.progress_bar.pack(fill="x", padx=10, pady=(0, 5))

        self.status_label = ctk.CTkLabel(
            self.progress_card,
            text="図面を選択して「解析を実行」を押してください",
            font=ctk.CTkFont(size=11),
            text_color=MUTED_TEXT,
            justify="left",
            anchor="w",
            wraplength=260,
        )
        self.status_label.pack(fill="x", padx=10, pady=(0, 7))

        # 右側パネル
        self.right_frame = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=CARD_BG,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        self.right_frame.grid(row=0, column=1, padx=(0, 20), pady=20, sticky="nsew")
        self.right_frame.grid_columnconfigure(0, weight=1)
        self.right_frame.grid_rowconfigure(2, weight=1)

        # プレビューコンテナ
        self.preview_container = ctk.CTkFrame(
            self.right_frame, fg_color="#E9ECEF", corner_radius=10, height=250
        )
        self.preview_container.grid(row=0, column=0, padx=20, pady=(14, 8), sticky="nsew")
        self.preview_container.grid_columnconfigure(0, weight=1)
        self.preview_container.grid_rowconfigure(0, weight=1)

        self.preview_label = ctk.CTkLabel(
            self.preview_container,
            text="図面プレビュー（ファイル選択で表示・クリックで拡大）",
            fg_color="transparent",
            text_color=MUTED_TEXT,
            height=250,
            cursor="hand2",
        )
        self.preview_label.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.preview_label.bind("<Button-1>", lambda e: self.open_image_viewer())

        self.zoom_btn = ctk.CTkButton(
            self.preview_container,
            text="🔍 拡大表示",
            width=96,
            height=28,
            fg_color=CARD_BG,
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            command=self.open_image_viewer,
            state="disabled",
        )
        self.zoom_btn.place(relx=0.985, rely=0.05, anchor="ne")

        # 製品最大外形寸法（最大長さ・最大幅・最大高さ）常時表示バー
        self.dims_bar = ctk.CTkFrame(
            self.right_frame,
            fg_color="#F0F6FF",
            corner_radius=8,
            border_width=1,
            border_color="#CFE2FF",
        )
        self.dims_bar.grid(row=1, column=0, padx=20, pady=(0, 6), sticky="ew")

        self.dims_bar_label = ctk.CTkLabel(
            self.dims_bar,
            text="📏 最大外形寸法:  最大長さ (L) --- mm  ｜  最大幅 (W) --- mm  ｜  最大高さ (H) --- mm",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
            anchor="w",
        )
        self.dims_bar_label.pack(side="left", padx=12, pady=6)

        # Treeview 共通ライトモードスタイル
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(
            "Treeview",
            background="white",
            foreground=TEXT_COLOR,
            fieldbackground="white",
            rowheight=24,
        )
        style.configure(
            "Treeview.Heading",
            background="#E9ECEF",
            foreground=TEXT_COLOR,
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Treeview", background=[("selected", "#CFE2FF")], foreground=[("selected", TEXT_COLOR)]
        )

        # 結果表示タブビュー (5タブ構成)
        self.result_tabview = ctk.CTkTabview(
            self.right_frame,
            fg_color=CARD_BG,
            segmented_button_fg_color="#DEE2E6",
            segmented_button_selected_color=PRIMARY,
            segmented_button_selected_hover_color=PRIMARY_HOVER,
            segmented_button_unselected_color="#6C757D",
            segmented_button_unselected_hover_color="#5C636A",
            text_color="white",
            border_width=1,
            border_color=BORDER_COLOR,
        )
        self.result_tabview.grid(row=2, column=0, padx=20, pady=(0, 8), sticky="nsew")

        tab_vol = self.result_tabview.add(TAB_VOLUME)
        tab_mat = self.result_tabview.add(TAB_MATERIALS)
        tab_pur = self.result_tabview.add(TAB_PURCHASED)
        tab_mac = self.result_tabview.add(TAB_MACHINED)
        tab_sum = self.result_tabview.add(TAB_SUMMARY)

        # タブ1: 体積・重量（数量・単重・小計重量対応）
        self.tree = self._create_result_treeview(
            tab_vol,
            [
                ("part_number", "品番", 50, "w"),
                ("part_name", "部品名/要素名", 145, "w"),
                ("material", "材質 (比重)", 100, "w"),
                ("quantity", "数量", 45, "e"),
                ("dimensions", "読み取り寸法", 135, "w"),
                ("volume", "体積 (mm³)", 95, "e"),
                ("unit_weight", "単重 (kg)", 80, "e"),
                ("weight", "小計重量 (kg)", 90, "e"),
                ("notes", "備考 (計算式等)", 220, "w"),
            ],
            sortable=False,
        )

        # タブ2: 部材表 (切断寸法)
        self.materials_tree = self._create_result_treeview(
            tab_mat,
            [
                ("part_number", "品番", 50, "w"),
                ("part_name", "部品名・部位", 120, "w"),
                ("material_type", "部材種別", 95, "w"),
                ("material_spec", "材質", 80, "w"),
                ("section_size", "断面規格・板厚", 110, "w"),
                ("cut_dimensions", "切断寸法 (mm)", 115, "w"),
                ("cut_angle_end", "端部切断", 90, "w"),
                ("quantity", "数量", 45, "e"),
                ("total_length", "合計長 (mm)", 85, "e"),
                ("weight", "小計重量 (kg)", 85, "e"),
                ("surface_treatment", "表面処理", 85, "w"),
                ("notes", "備考", 140, "w"),
            ],
            sortable=True,
        )

        # タブ3: 使用購入品 (手配)
        self.purchased_tree = self._create_result_treeview(
            tab_pur,
            [
                ("part_number", "品番", 50, "w"),
                ("part_name", "品名", 135, "w"),
                ("manufacturer", "メーカー", 95, "w"),
                ("model_number", "型式・規格", 135, "w"),
                ("material_or_spec", "材質・仕様", 115, "w"),
                ("quantity", "数量", 50, "e"),
                ("unit", "単位", 45, "w"),
                ("procurement_category", "手配区分", 95, "w"),
                ("notes", "備考", 160, "w"),
            ],
            sortable=True,
        )

        # タブ4: 使用機械加工品
        self.machined_tree = self._create_result_treeview(
            tab_mac,
            [
                ("part_number", "品番", 50, "w"),
                ("part_name", "部品名", 120, "w"),
                ("material", "材質", 80, "w"),
                ("blank_or_size", "素材・概略寸法", 120, "w"),
                ("machining_processes", "加工工程・内容", 190, "w"),
                ("precision_surface", "公差・粗さ", 115, "w"),
                ("heat_surface", "熱処理・表面処理", 115, "w"),
                ("quantity", "数量", 45, "e"),
                ("notes", "備考", 135, "w"),
            ],
            sortable=True,
        )

        # タブ5: 見積り・搬送サマリー
        self._build_summary_tab(tab_sum)

        # 下部フッター
        footer_row = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        footer_row.grid(row=3, column=0, padx=20, pady=(0, 14), sticky="ew")
        ctk.CTkLabel(
            footer_row,
            text="※右クリックで行/表コピー可（部材・購入品・加工品・集計は列見出しクリックで並べ替え）",
            font=ctk.CTkFont(size=11),
            text_color=MUTED_TEXT,
        ).pack(side="left")

        self.total_label = ctk.CTkLabel(
            footer_row,
            text="全体合計: --- kg / 概算 --- 円",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=TEXT_COLOR,
        )
        self.total_label.pack(side="right")

    def select_file(self):
        path = filedialog.askopenfilename(
            title="図面ファイルを選択", filetypes=[("PDF/Image", "*.pdf *.png *.jpg *.jpeg *.bmp")]
        )
        if not path:
            return
        self.selected_file_path = path
        self.selected_page_index = 0
        self.pdf_page_count = get_pdf_page_count(path)
        self.file_path_display.configure(text=os.path.basename(path), text_color=TEXT_COLOR)

        page_labels = [f"{i + 1} / {self.pdf_page_count}" for i in range(self.pdf_page_count)]
        self._set_page_selector(
            page_labels, page_labels[0], enabled=(self.pdf_page_count > 1)
        )
        self._load_preview_async(path, 0)

    def _on_page_selected(self, selected_label):
        try:
            idx = int(str(selected_label).split("/")[0].strip()) - 1
            idx = max(0, min(idx, self.pdf_page_count - 1))
        except Exception:
            idx = 0
        self.selected_page_index = idx
        if self.selected_file_path:
            self._load_preview_async(self.selected_file_path, idx)

    def _load_preview_async(self, path, page_idx):
        self.update_status("図面プレビューを読み込み中...")

        def _worker():
            try:
                img = load_preview_image(path, page_index=page_idx)
                self.update_preview(img)
                self.update_status("図面読み込み完了（解析を実行できます）")
            except Exception as e:
                self.update_status(f"プレビュー読み込みエラー: {e}")

        threading.Thread(target=_worker, daemon=True).start()

    def open_image_viewer(self):
        if self.processed_image_pil is None:
            return
        fname = os.path.basename(self.selected_file_path) if self.selected_file_path else "図面"
        page_note = (
            f" (P.{self.selected_page_index + 1}/{self.pdf_page_count})"
            if self.pdf_page_count > 1
            else ""
        )
        ImageViewerDialog(
            self, self.processed_image_pil, title=f"図面拡大プレビュー - {fname}{page_note}"
        )

    def open_material_price_dialog(self):
        extracted_mats = []
        if self.last_result_data:
            for p in self.last_result_data.get("parts", []):
                m = str(p.get("material") or "").strip()
                if m and m not in extracted_mats:
                    extracted_mats.append(m)
            for item in self.last_result_data.get("materials", []):
                m = str(item.get("material_spec") or "").strip()
                if m and m not in extracted_mats:
                    extracted_mats.append(m)

        cur_default = self._get_effective_default_unit_price()

        def _on_apply(new_prices, new_default):
            self.api_settings["material_unit_prices"].update(new_prices)
            self.api_settings["default_unit_price"] = new_default
            self.unit_price_entry.delete(0, "end")
            self.unit_price_entry.insert(0, str(int(round(new_default))))
            self.save_settings()
            if self.last_result_data:
                self._refresh_summary_views(self.last_result_data)
                self.update_status("新しい単価で見積り金額を再計算しました")

        MaterialPriceDialog(
            self,
            extracted_mats,
            self.api_settings.get("material_unit_prices", {}),
            cur_default,
            _on_apply,
        )

    def update_status(self, text):
        self._safe_after(0, lambda t=text: self.status_label.configure(text=t))

    @staticmethod
    def _format_elapsed(seconds):
        sec_int = max(0, int(seconds))
        mins, secs = divmod(sec_int, 60)
        return f"{mins:02d}:{secs:02d}"

    def _refresh_progress_meta(self):
        pct = int(round(max(0.0, min(1.0, self._current_progress_ratio)) * 100))
        if self._analysis_start_time is not None:
            elapsed_str = self._format_elapsed(time.perf_counter() - self._analysis_start_time)
            meta_txt = f"{pct}%  |  {elapsed_str}"
        else:
            meta_txt = f"{pct}%"
        self.progress_meta_label.configure(text=meta_txt)
        if self._is_analyzing and self._analysis_start_time is not None:
            elapsed_str = self._format_elapsed(time.perf_counter() - self._analysis_start_time)
            self.total_label.configure(text=f"解析中... {pct}%（経過 {elapsed_str}）")

    def _tick_elapsed_timer(self):
        self._elapsed_timer_id = None
        if not self._is_analyzing or not self.winfo_exists():
            return
        self._refresh_progress_meta()
        try:
            self._elapsed_timer_id = self.after(500, self._tick_elapsed_timer)
        except Exception:
            self._elapsed_timer_id = None

    def _start_elapsed_timer(self):
        self._stop_elapsed_timer()
        self._analysis_start_time = time.perf_counter()
        self._refresh_progress_meta()
        try:
            self._elapsed_timer_id = self.after(500, self._tick_elapsed_timer)
        except Exception:
            self._elapsed_timer_id = None

    def _stop_elapsed_timer(self):
        if self._elapsed_timer_id is not None:
            try:
                self.after_cancel(self._elapsed_timer_id)
            except Exception:
                pass
            self._elapsed_timer_id = None

    def _apply_progress_ui(self, ratio, phase="", detail="", bar_color=None):
        """UIスレッド上で進捗バー・フェーズ見出し・進捗率・詳細メッセージを直接更新する"""
        clamped = max(0.0, min(1.0, float(ratio)))
        self._current_progress_ratio = clamped
        self.progress_bar.set(clamped)
        if bar_color is not None:
            self.progress_bar.configure(progress_color=bar_color)
        if phase:
            self.progress_phase_label.configure(text=phase)
        if detail:
            self.status_label.configure(text=detail)
        self._refresh_progress_meta()

    def update_progress(self, ratio, phase="", detail="", bar_color=None):
        """解析実行中の進捗をスレッドセーフに更新する（完了・中止後の遅延通知は無視）"""
        clamped = max(0.0, min(1.0, float(ratio)))

        def _apply():
            if not self._is_analyzing:
                return
            self._apply_progress_ui(clamped, phase=phase, detail=detail, bar_color=bar_color)

        self._safe_after(0, _apply)

    def open_api_settings(self):
        APISettingsDialog(self, self.api_settings, self.on_api_settings_saved)

    def open_help(self, initial_tab="guide"):
        HelpDialog(
            self,
            on_open_api_settings=self.open_api_settings,
            settings_dir=SETTINGS_DIR,
            initial_tab=initial_tab,
        )

    def on_api_settings_saved(self, new_settings):
        new_settings["analysis_targets"] = self._get_selected_targets()
        new_settings["default_density"] = self.api_settings.get("default_density", 7.85)
        new_settings["default_unit_price"] = self.api_settings.get("default_unit_price", 250.0)
        new_settings["material_unit_prices"] = copy.deepcopy(
            self.api_settings.get("material_unit_prices", DEFAULT_MATERIAL_UNIT_PRICES)
        )
        self.api_settings = new_settings
        self._refresh_api_panel()
        self.save_settings()

    def update_preview(self, pil_image):
        self.processed_image_pil = pil_image
        temp_img = pil_image.copy()
        temp_img.thumbnail((680, 248), Image.Resampling.LANCZOS)
        ctk_img = ctk.CTkImage(light_image=temp_img, dark_image=temp_img, size=temp_img.size)

        def _apply():
            self.preview_label.configure(image=ctk_img, text="")
            self.zoom_btn.configure(state="normal")

        self._safe_after(0, _apply)

    def _clear_all_result_trees(self):
        for tv in (
            self.tree,
            self.materials_tree,
            self.purchased_tree,
            self.machined_tree,
            self.cost_summary_tree,
            self.stock_summary_tree,
        ):
            tv.delete(*tv.get_children())
        if hasattr(self, "dims_bar_label"):
            self.dims_bar_label.configure(
                text="📏 最大外形寸法:  最大長さ (L) --- mm  ｜  最大幅 (W) --- mm  ｜  最大高さ (H) --- mm"
            )
        if hasattr(self, "dim_length_label"):
            self.dim_length_label.configure(text="--- mm")
            self.dim_width_label.configure(text="--- mm")
            self.dim_height_label.configure(text="--- mm")
            self.dim_weight_label.configure(text="--- kg")
        if hasattr(self, "truck_badge"):
            self.truck_badge.configure(text=" 解析後に自動判定 ")

    def _set_analyzing_state(self, active):
        self._is_analyzing = active
        if active:
            self.analyze_btn.configure(
                text="⏹ 解析を中止",
                fg_color=COLOR_DANGER,
                hover_color=COLOR_DANGER_HOVER,
                state="normal",
            )
            self.select_btn.configure(state="disabled")
            self.export_btn.configure(state="disabled")
            self._current_progress_ratio = 0.02
            self.progress_bar.configure(progress_color=PRIMARY)
            self.progress_bar.set(0.02)
            self.progress_phase_label.configure(text="解析開始中...", text_color=PRIMARY)
            self._start_elapsed_timer()
        else:
            self._stop_elapsed_timer()
            self.analyze_btn.configure(
                text="解析を実行",
                fg_color=COLOR_SUCCESS,
                hover_color=COLOR_SUCCESS_HOVER,
                state="normal",
            )
            self.select_btn.configure(state="normal")

    def on_analyze_or_cancel_click(self):
        if self._is_analyzing:
            if self._cancel_event is not None and not self._cancel_event.is_set():
                self._cancel_event.set()
                self.analyze_btn.configure(text="中止処理中...", state="disabled")
                self.update_status("解析の中止をリクエストしました...")
            return
        self.start_analysis_thread()

    def start_analysis_thread(self):
        if not self.selected_file_path:
            messagebox.showwarning("警告", "ファイルを選択してください。", parent=self)
            return

        targets = self._get_selected_targets()
        if not any(targets.values()):
            messagebox.showwarning(
                "警告", "「解析項目の選択」で1つ以上の項目にチェックを入れてください。", parent=self
            )
            return

        engine = self.api_settings.get("engine", "gemini")

        if engine == "azure":
            az_key = self.api_settings.get("azure_api_key", "").strip()
            az_ep_raw = self.api_settings.get("azure_endpoint", "").strip()
            az_dep1 = self.api_settings.get("azure_deployment_step1", "").strip()
            az_ver = self.api_settings.get("azure_api_version", DEFAULT_AZURE_API_VERSION).strip()
            base_ep, resolved_dep, resolved_ver = parse_azure_endpoint_info(
                az_ep_raw, az_dep1, az_ver
            )
            if base_ep:
                self.api_settings["azure_endpoint"] = base_ep
            if resolved_dep and not az_dep1:
                self.api_settings["azure_deployment_step1"] = resolved_dep
            if resolved_ver:
                self.api_settings["azure_api_version"] = resolved_ver

            if not az_key or not base_ep or not resolved_dep:
                if messagebox.askyesno(
                    "接続設定が不足しています",
                    "Azure OpenAI Service の接続情報（エンドポイントURL、APIキー、デプロイ名）が不足しています。\n\n今すぐ「⚙️ AI詳細設定」画面を開きますか？",
                    parent=self,
                ):
                    self.open_api_settings()
                return
        else:
            plan = self.api_settings.get("plan", "free")
            plan_ja = "無料枠" if plan == "free" else "課金枠"
            if not self.api_settings.get(f"{plan}_key", "").strip():
                if messagebox.askyesno(
                    "APIキーが未設定です",
                    f"Gemini API（{plan_ja}）のAPIキーが設定されていません。\n\n今すぐ「⚙️ AI詳細設定」画面を開きますか？",
                    parent=self,
                ):
                    self.open_api_settings()
                return

        if not self.save_basic_params(explicit=True):
            return
        self._cancel_event = threading.Event()
        self._set_analyzing_state(True)
        self._clear_all_result_trees()
        self.total_label.configure(text="解析中... 2%（経過 00:00）")
        settings_snapshot = copy.deepcopy(self.api_settings)
        density_val = self._format_density(self.api_settings.get("default_density", 7.85))
        file_path = self.selected_file_path
        page_idx = self.selected_page_index
        cancel_ev = self._cancel_event
        threading.Thread(
            target=self.run_analysis,
            args=(file_path, density_val, settings_snapshot, targets, page_idx, cancel_ev),
            daemon=True,
        ).start()

    def _parse_json_response(self, result_text):
        return parse_json_response(result_text, status_callback=self.update_status)

    def _repair_and_parse_json(self, json_str):
        return repair_and_parse_json(json_str)

    def _safe_float(self, val):
        return safe_float(val)

    def _evaluate_formula(self, formula_str):
        return evaluate_formula(formula_str)

    def run_analysis(
        self, file_path, density, settings, targets=None, page_index=0, cancel_event=None
    ):
        engine = settings.get("engine", "gemini")
        if targets is None:
            targets = {
                "volume": True,
                "materials": False,
                "purchased_parts": False,
                "machined_parts": False,
            }
        start_time = time.perf_counter()
        try:
            result_data = run_drawing_analysis(
                file_path=file_path,
                density=density,
                settings=settings,
                targets=targets,
                status_callback=self.update_status,
                preview_callback=self.update_preview,
                page_index=page_index,
                cancel_event=cancel_event,
                progress_callback=self.update_progress,
            )
            elapsed = time.perf_counter() - start_time

            def _finish_success(d=result_data, el=elapsed):
                # 1. 先に解析中フラグを解除し経過タイマーを停止（total_label上書きを防止）
                self._set_analyzing_state(False)
                # 2. 進捗カードを100%完了表示へ更新
                self._apply_progress_ui(
                    1.0,
                    phase="✅ 解析完了",
                    detail=f"解析完了 (所要時間: {el:.1f}秒)",
                    bar_color=COLOR_SUCCESS,
                )
                # 3. 結果テーブル・サマリー・下部総重量/材料費ラベルを描画
                self.display_results(d)

            self._safe_after(0, _finish_success)

        except AnalysisCancelledError:
            def _finish_cancelled():
                self._set_analyzing_state(False)
                self._apply_progress_ui(
                    self._current_progress_ratio,
                    phase="⏹ 解析を中止しました",
                    detail="ユーザー操作により解析を中止しました",
                    bar_color="#ADB5BD",
                )
                self.total_label.configure(text="解析が中止されました")

            self._safe_after(0, _finish_cancelled)
        except DiagnosedAPIError as api_err:
            def _finish_api_error(d=api_err.diag):
                self._set_analyzing_state(False)
                self._apply_progress_ui(
                    self._current_progress_ratio,
                    phase="⚠️ 通信・APIエラー",
                    detail="通信・APIエラーが発生しました",
                    bar_color=COLOR_DANGER,
                )
                self.total_label.configure(text="全体合計: --- kg / 概算 --- 円")
                show_api_error_dialog(self, d)

            self._safe_after(0, _finish_api_error)
        except Exception as e:
            diag = diagnose_api_error(
                "azure" if engine == "azure" else "gemini",
                e,
                endpoint=settings.get("azure_endpoint", "") if engine == "azure" else "",
            )

            def _finish_error(d=diag):
                self._set_analyzing_state(False)
                self._apply_progress_ui(
                    self._current_progress_ratio,
                    phase="⚠️ エラー発生",
                    detail="解析中にエラーが発生しました",
                    bar_color=COLOR_DANGER,
                )
                self.total_label.configure(text="全体合計: --- kg / 概算 --- 円")
                show_api_error_dialog(self, d)

            self._safe_after(0, _finish_error)

    def _refresh_summary_views(self, data):
        """見積り・搬送サマリータブ・最大外形寸法バー・下部合計ラベルを最新単価で再描画する"""
        default_price = self._get_effective_default_unit_price()
        default_den = safe_float(data.get("default_density", 7.85)) or 7.85
        summary = build_estimation_and_logistics_summary(
            data,
            custom_prices=self.api_settings.get("material_unit_prices", {}),
            default_unit_price=default_price,
            default_density=default_den,
        )
        if not summary:
            return

        # 製品最大外形寸法（最大長さ・最大幅・最大高さ）＆ 総重量の更新
        truck = summary["truck"]
        od = summary["overall_dimensions"]
        len_txt = f"{od['length_mm']:,.0f} mm" if od["length_mm"] > 0 else "--- mm"
        wid_txt = f"{od['width_mm']:,.0f} mm" if od["width_mm"] > 0 else "--- mm"
        hei_txt = f"{od['height_mm']:,.0f} mm" if od["height_mm"] > 0 else "--- mm"
        wt_txt = f"{summary['total_weight_kg']:,.2f} kg" if summary["total_weight_kg"] > 0 else "--- kg"

        self.dim_length_label.configure(text=len_txt)
        self.dim_width_label.configure(text=wid_txt)
        self.dim_height_label.configure(text=hei_txt)
        self.dim_weight_label.configure(text=wt_txt)
        self.dims_bar_label.configure(
            text=f"📏 最大外形寸法:  最大長さ (L) {len_txt}  ｜  最大幅 (W) {wid_txt}  ｜  最大高さ (H) {hei_txt}"
        )

        # トラック選定カード更新
        self.truck_badge.configure(text=f" 推奨: {truck['recommended_truck']} ")

        detail_lines = [
            f"・荷台/積載目安: {truck['bed_spec']}",
            f"・判定根拠: {truck['reasons']}",
            f"・荷役アドバイス: {truck['handling_advice']}",
        ]
        if od.get("transport_notes"):
            detail_lines.append(f"・図面上の搬送注記: {od['transport_notes']}")
        self.truck_detail_label.configure(text="\n".join(detail_lines))

        # 材質別重量＆概算材料費テーブル更新
        self.cost_summary_tree.delete(*self.cost_summary_tree.get_children())
        for mc in summary.get("material_cost_rows", []):
            self.cost_summary_tree.insert(
                "",
                "end",
                values=(
                    mc["material"],
                    f"{mc['density_g_cm3']:.2f}",
                    f"{mc['weight_kg']:,.2f}",
                    f"{mc['unit_price_jpy']:,.0f}",
                    f"{int(round(mc['estimated_cost_jpy'])):,}",
                ),
            )
        self.cost_total_label.configure(
            text=f"総重量: {summary['total_weight_kg']:,.2f} kg ／ 概算材料費: {int(round(summary['total_material_cost_jpy'])):,} 円"
        )

        # 鋼材定尺発注集計テーブル更新
        self.stock_summary_tree.delete(*self.stock_summary_tree.get_children())
        for st_row in summary.get("stock_summary_rows", []):
            self.stock_summary_tree.insert(
                "",
                "end",
                values=(
                    st_row["material_type"],
                    st_row["material_spec"],
                    st_row["section_size"],
                    str(st_row["total_pieces"]),
                    f"{st_row['total_length_mm']:,.0f}" if st_row["total_length_mm"] > 0 else "-",
                    st_row["stock_guide"],
                    f"{st_row['total_weight_kg']:,.2f}" if st_row["total_weight_kg"] > 0 else "-",
                ),
            )

        # フッターのコンパクトサマリー更新
        targets = data.get("targets", {"volume": True})
        summary_parts = []
        if summary["total_weight_kg"] > 0:
            summary_parts.append(
                f"総重量: {summary['total_weight_kg']:,.2f} kg / 概算材料費: {int(round(summary['total_material_cost_jpy'])):,} 円 / {truck['recommended_truck']}"
            )
        counts = []
        if targets.get("materials"):
            counts.append(f"部材: {len(data.get('materials', []))}件")
        if targets.get("purchased_parts"):
            counts.append(f"購入品: {len(data.get('purchased_parts', []))}件")
        if targets.get("machined_parts"):
            counts.append(f"加工品: {len(data.get('machined_parts', []))}件")
        if counts:
            summary_parts.append(" / ".join(counts))

        self.total_label.configure(
            text="  ｜  ".join(summary_parts) if summary_parts else "全体合計: --- kg / 概算 --- 円"
        )

    def display_results(self, data):
        """解析結果を全5タブの Treeview へ表示し、見積り・搬送サマリーを更新する"""
        self._clear_all_result_trees()
        targets = data.get("targets", {"volume": True})
        default_den = safe_float(data.get("default_density", 7.85)) or 7.85

        # 1. 体積・重量タブ
        parts = data.get("parts", [])
        calculate_volume_and_weight(parts, default_density=default_den)

        for part in parts:
            p_num = str(part.get("part_number") or "")
            p_name = str(part.get("part_name") or "不明な部品")
            p_mat = str(part.get("material") or "")
            p_den = safe_float(part.get("density_g_cm3"))
            p_qty = safe_quantity(part.get("quantity", 1))
            unit_wt = safe_float(part.get("unit_weight_kg", 0.0))
            sub_vol = safe_float(part.get("calculated_volume_mm3", 0.0))
            sub_wt = safe_float(part.get("calculated_weight_kg", 0.0))

            if p_mat and p_den > 0.0:
                mat_display = f"{p_mat} ({p_den:.2f})"
            elif p_den > 0.0:
                mat_display = f"不明 ({p_den:.2f})"
            elif p_mat:
                mat_display = p_mat
            else:
                mat_display = "---"

            p_id = self.tree.insert(
                "",
                "end",
                values=(
                    p_num,
                    p_name,
                    mat_display,
                    str(p_qty),
                    "",
                    f"{sub_vol:,.1f}" if sub_vol != 0.0 else "---",
                    f"{unit_wt:,.2f}" if unit_wt != 0.0 else "---",
                    f"{sub_wt:,.2f}" if sub_wt != 0.0 else "---",
                    "",
                ),
                open=True,
            )

            for elem in part.get("elements", []):
                e_name = str(elem.get("element_name") or "要素")
                e_dim = str(elem.get("dimensions") or "")
                e_calc = str(elem.get("calculation_formula") or "")
                e_vol = safe_float(elem.get("calculated_volume_mm3", 0.0))
                e_wt = safe_float(elem.get("calculated_weight_kg", 0.0))
                e_notes = str(elem.get("notes") or "")

                combined_notes = f"[式: {e_calc}] {e_notes}".strip() if e_calc else e_notes
                self.tree.insert(
                    p_id,
                    "end",
                    values=(
                        "",
                        f"  ├ {e_name}",
                        "",
                        "",
                        e_dim,
                        f"{e_vol:,.1f}" if e_vol != 0.0 else "---",
                        f"{e_wt:,.3f}" if e_wt != 0.0 else "---",
                        f"{(e_wt * p_qty):,.3f}" if e_wt != 0.0 else "---",
                        combined_notes,
                    ),
                )

        # 2. 部材表 (切断寸法) タブ
        materials = data.get("materials", [])
        calculate_materials_metrics(materials, default_density=default_den)
        for item in materials:
            tot_l = safe_float(item.get("total_length_mm", 0.0))
            tot_wt = safe_float(item.get("total_weight_kg", 0.0))
            self.materials_tree.insert(
                "",
                "end",
                values=(
                    str(item.get("part_number") or ""),
                    str(item.get("part_name") or ""),
                    str(item.get("material_type") or ""),
                    str(item.get("material_spec") or ""),
                    str(item.get("section_size") or ""),
                    str(item.get("cut_dimensions") or ""),
                    str(item.get("cut_angle_end") or ""),
                    str(safe_quantity(item.get("quantity", 1))),
                    f"{tot_l:,.0f}" if tot_l > 0 else "-",
                    f"{tot_wt:,.2f}" if tot_wt > 0 else "-",
                    str(item.get("surface_treatment") or ""),
                    str(item.get("notes") or ""),
                ),
            )

        # 3. 使用購入品 (手配) タブ
        purchased = data.get("purchased_parts", [])
        for item in purchased:
            self.purchased_tree.insert(
                "",
                "end",
                values=(
                    str(item.get("part_number") or ""),
                    str(item.get("part_name") or ""),
                    str(item.get("manufacturer") or ""),
                    str(item.get("model_number") or ""),
                    str(item.get("material_or_spec") or ""),
                    str(safe_quantity(item.get("quantity", 1))),
                    str(item.get("unit") or "個"),
                    str(item.get("procurement_category") or ""),
                    str(item.get("notes") or ""),
                ),
            )

        # 4. 使用機械加工品タブ
        machined = data.get("machined_parts", [])
        for item in machined:
            self.machined_tree.insert(
                "",
                "end",
                values=(
                    str(item.get("part_number") or ""),
                    str(item.get("part_name") or ""),
                    str(item.get("material") or ""),
                    str(item.get("blank_or_size") or ""),
                    str(item.get("machining_processes") or ""),
                    str(item.get("precision_surface") or ""),
                    str(item.get("heat_surface_treatment") or ""),
                    str(safe_quantity(item.get("quantity", 1))),
                    str(item.get("notes") or ""),
                ),
            )

        # 5. 見積り・搬送サマリータブ＆フッター更新
        self._refresh_summary_views(data)

        tab_priority = [
            ("volume", TAB_VOLUME),
            ("materials", TAB_MATERIALS),
            ("purchased_parts", TAB_PURCHASED),
            ("machined_parts", TAB_MACHINED),
        ]
        for key, tab_name in tab_priority:
            if targets.get(key):
                self.result_tabview.set(tab_name)
                break

        self.last_result_data = data
        default_price = self._get_effective_default_unit_price()
        has_any_data = bool(
            build_export_tables(
                data,
                custom_prices=self.api_settings.get("material_unit_prices", {}),
                default_unit_price=default_price,
            )
        )
        self.export_btn.configure(state="normal" if has_any_data else "disabled")

    def open_in_excel(self):
        """保持している解析結果を印刷設定済みの新規Excelブックとして開く（必要時に任意保存）"""
        if not self.last_result_data:
            return

        default_price = self._get_effective_default_unit_price()
        tables = build_export_tables(
            self.last_result_data,
            custom_prices=self.api_settings.get("material_unit_prices", {}),
            default_unit_price=default_price,
        )
        if not tables:
            messagebox.showwarning("警告", "Excelで開くデータがありません。", parent=self)
            return

        page_note = (
            f" (P.{self.selected_page_index + 1}/{self.pdf_page_count})"
            if self.pdf_page_count > 1
            else ""
        )

        self.update_status("印刷設定済みの新規Excelブックを開いています...")
        try:
            open_tables_in_new_excel(
                tables,
                source_file_path=self.selected_file_path or "",
                page_note=page_note,
            )
            self.update_status(
                "新規Excelブックで開きました（A4横印刷設定済み・必要時に任意保存できます）"
            )
        except Exception as e:
            self.update_status("Excelの自動起動に失敗しました。")
            if messagebox.askyesno(
                "Excelを起動できませんでした",
                f"Excelの自動起動中にエラーが発生しました。\n"
                f"（Excelが未インストールの可能性があります）\n\n"
                f"詳細: {str(e)}\n\n"
                f"代わりに印刷設定済みのExcelファイル (.xlsx) として保存しますか？",
                parent=self,
            ):
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel Files", "*.xlsx")],
                    title="保存先を選択",
                )
                if file_path:
                    try:
                        save_tables_as_xlsx(
                            tables,
                            file_path,
                            source_file_path=self.selected_file_path or "",
                            page_note=page_note,
                        )
                        messagebox.showinfo(
                            "保存完了", f"Excelファイルを保存しました。\n{file_path}", parent=self
                        )
                    except Exception as save_err:
                        messagebox.showerror(
                            "保存エラー",
                            f"ファイルの保存中にエラーが発生しました。\n詳細:\n{str(save_err)}",
                            parent=self,
                        )


if __name__ == "__main__":
    app = DrawingAnalysisApp()
    app.mainloop()
