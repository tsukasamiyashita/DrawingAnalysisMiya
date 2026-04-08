import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
import pdfplumber # ベクターテキスト解析用
import cv2
import numpy as np
from PIL import Image, ImageTk
import io
import os
import json
import threading
import concurrent.futures
import time
from pydantic import BaseModel, Field
from typing import List
import google.generativeai as genai
import re
import pandas as pd # Excel出力用

# 分離したモジュールを読み込む
from settings_dialog import APISettingsDialog

# --- 設定ファイルの保存先定義 ---
SETTINGS_DIR = os.path.join(os.path.expanduser("~"), "DrawingAnalysisMiya")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")

# --- 構造化データ（Structured Output）のスキーマ定義 (マルチターン用) ---

class Element(BaseModel):
    element_name: str = Field(description="構成要素の名称（例: ベースプレート、リブ、ボス、貫通穴）")
    dimensions: str = Field(description="三面図（正面図、平面図、側面図など）の各投影ビューから読み取った寸法の統合情報。絶対に空欄にせず、立体の幅・高さ・奥行きがどう定義されているか記載すること（例: 正面図W150xH100, 平面図D10）")
    calculation_formula: str = Field(description="三面図から復元した3次元形状に基づく、体積(mm³)を導き出すための『純粋な数式のみ』。文字や単位を含めず、Pythonのeval関数で計算可能な文字列にすること。穴などの空間は必ずマイナス（引き算）として式に組み込むこと。例: '150 * 100 * 10' または '-(50/2)**2 * 3.14159 * 10'")
    notes: str = Field(description="三面図から立体形状をどう解釈したか等の特記事項がない場合は必ず空文字にすること。記述する場合も極力短く（15文字以内）し、出力トークンを節約すること。")

# ステップ1用: 部品の基本情報のみ
class PartBasic(BaseModel):
    part_number: str = Field(description="部品番号（バルーン記号や部品表の番号）。ない場合は空文字にする")
    part_name: str = Field(description="部品名（部品表に記載されている名称をそのまま転記）。不明な場合は '不明な部品' とする")
    material: str = Field(description="部品の材質（図面から読み取れる場合。例: SS400, SUS304, AL, 樹脂 など）。ない場合は空文字にする")
    density_g_cm3: float = Field(description="材質に基づいた比重（g/cm³）。図面に記載がない場合はデフォルトの比重を使用すること。不明な場合は 0.0 とする")

class PartListResult(BaseModel):
    parts: List[PartBasic] = Field(description="図面から抽出されたすべての部品の基本情報のリスト")

# ステップ2用: 指定部品の要素リスト
class PartElementsResult(BaseModel):
    elements: List[Element] = Field(description="指定された部品を構成する要素のリスト")

# ステップ3用: 最終的な完全な部品情報
class CompletePart(BaseModel):
    part_number: str = Field(description="部品番号")
    part_name: str = Field(description="部品名")
    material: str = Field(description="材質")
    density_g_cm3: float = Field(description="比重")
    elements: List[Element] = Field(description="この部品を構成する要素のリスト")

class MissingPartsResult(BaseModel):
    missing_parts: List[CompletePart] = Field(description="検証フェーズで新たに発見された、抽出漏れの部品とその要素のリスト。見落としがない場合は空リストにする。")


# --- メインアプリケーションクラス ---
class DrawingAnalysisApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("DrawingAnalysisMiya - 2D機械図面 体積解析ツール")
        self.geometry("1100x750")
        self.minsize(900, 650)
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.selected_file_path = None
        self.processed_image_pil = None
        self.last_result_data = None # Excel出力用にデータを保持
        
        # モデル別パフォーマンス(RPM/スレッド)のデフォルト値
        self.default_free_perf = {
            "gemini-3-flash": {"rpm": 15, "threads": 1},
            "gemini-3.1-pro-preview": {"rpm": 2, "threads": 1},
            "gemini-3.1-flash-lite-preview": {"rpm": 15, "threads": 1},
            "gemini-2.5-flash": {"rpm": 15, "threads": 1},
            "gemini-2.5-pro": {"rpm": 2, "threads": 1}
        }
        self.default_paid_perf = {
            "gemini-3-flash": {"rpm": 300, "threads": 5},
            "gemini-3.1-pro-preview": {"rpm": 150, "threads": 5},
            "gemini-3.1-flash-lite-preview": {"rpm": 300, "threads": 5},
            "gemini-2.5-flash": {"rpm": 1000, "threads": 5},
            "gemini-2.5-pro": {"rpm": 360, "threads": 5}
        }
        
        # API設定を保持する辞書
        self.api_settings = {
            "plan": "free",
            "free_key": "",
            "paid_key": "",
            
            "free_model_step1": "gemini-3-flash",
            "free_model_step2": "gemini-3-flash",
            "free_model_step3": "gemini-3.1-pro-preview",
            "paid_model_step1": "gemini-3-flash",
            "paid_model_step2": "gemini-3-flash",
            "paid_model_step3": "gemini-3.1-pro-preview",
            
            "free_model_perf": self.default_free_perf.copy(),
            "paid_model_perf": self.default_paid_perf.copy(),
            
            "free_temp": 0.0,
            "paid_temp": 0.0,
            "free_safety": True,
            "paid_safety": True,
            "free_tokens": 65535,
            "paid_tokens": 65535,
            "free_prompts": [],
            "paid_prompts": [],
            "saved_prompts": []
        }

        self.load_settings()
        self._setup_ui()
        self.after(200, self._maximize_window)

    def _maximize_window(self):
        self.update_idletasks()
        try:
            self.state('zoomed')
        except:
            pass

    def load_settings(self):
        """設定を読み込み、UIに反映させる。過去の設定からの自動移行も行う"""
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    
                    # 古い単一モデル・スレッド設定の移行処理
                    for plan in ["free", "paid"]:
                        old_model_key = f"{plan}_model"
                        if old_model_key in loaded:
                            old_val = loaded[old_model_key]
                            if "1.5" in old_val or "2.0" in old_val:
                                old_val = "gemini-3-flash"
                            for step in ["step1", "step2", "step3"]:
                                step_key = f"{plan}_model_{step}"
                                if step_key not in loaded:
                                    loaded[step_key] = old_val
                            del loaded[old_model_key]
                    
                    # 結合
                    self.api_settings.update(loaded)
                    
                    # 移行: 古い rpm / threads キーのみの時代のデータならデフォルトモデルパフォーマンスを補完
                    if "free_model_perf" not in loaded:
                        self.api_settings["free_model_perf"] = self.default_free_perf.copy()
                    if "paid_model_perf" not in loaded:
                        self.api_settings["paid_model_perf"] = self.default_paid_perf.copy()

                    # 古い構造の不要なキーを削除 (クリーンアップ)
                    keys_to_delete = [k for k in self.api_settings.keys() if ("rpm" in k or "threads" in k) and k not in ["free_model_perf", "paid_model_perf"]]
                    for k in keys_to_delete:
                        del self.api_settings[k]

                    # 1.5 などの古いモデルの強制変換
                    for plan in ["free", "paid"]:
                        for step in ["step1", "step2", "step3"]:
                            step_key = f"{plan}_model_{step}"
                            if step_key in self.api_settings and ("1.5" in self.api_settings[step_key] or "2.0" in self.api_settings[step_key]):
                                self.api_settings[step_key] = "gemini-3-flash"

                    if "models_list" in self.api_settings:
                        if any("1.5" in item[1] for item in self.api_settings["models_list"]):
                            del self.api_settings["models_list"]
                            
            except Exception as e:
                print(f"Load settings error: {e}")

    def save_settings(self):
        """設定を保存する"""
        try:
            os.makedirs(SETTINGS_DIR, exist_ok=True)
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.api_settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Save settings error: {e}")

    def _setup_ui(self):
        # 左側パネル
        self.left_frame = ctk.CTkFrame(self, width=320, corner_radius=10)
        self.left_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        # APIキー
        ctk.CTkLabel(self.left_frame, text="Gemini API キー:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=(20, 5))
        self.api_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.api_frame.pack(fill="x", padx=20)
        self.api_key_entry = ctk.CTkEntry(self.api_frame, show="*", placeholder_text="AIza...")
        self.api_key_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ctk.CTkButton(self.api_frame, text="⚙️", width=40, command=self.open_api_settings).pack(side="right")
        
        plan = self.api_settings.get("plan", "free")
        self.api_key_entry.insert(0, self.api_settings.get(f"{plan}_key", ""))

        self.plan_indicator = ctk.CTkLabel(self.left_frame, text="🟢 無料枠 (Free)" if plan=="free" else "🔵 課金枠 (Paid)", 
                                         text_color="#198754" if plan=="free" else "#0D6EFD", font=ctk.CTkFont(weight="bold"))
        self.plan_indicator.pack(anchor="w", padx=20, pady=(0, 20))

        # ファイル
        ctk.CTkLabel(self.left_frame, text="図面ファイル:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=(10, 5))
        self.select_btn = ctk.CTkButton(self.left_frame, text="ファイルを選択", command=self.select_file)
        self.select_btn.pack(fill="x", padx=20, pady=(0, 10))
        self.file_path_display = ctk.CTkLabel(self.left_frame, text="未選択", text_color="gray", wraplength=250)
        self.file_path_display.pack(anchor="w", padx=20, pady=(0, 20))

        # 比重
        ctk.CTkLabel(self.left_frame, text="基本比重 (g/cm³):\n(材質不明時に適用)", font=ctk.CTkFont(weight="bold"), justify="left").pack(anchor="w", padx=20, pady=(5, 5))
        self.density_entry = ctk.CTkEntry(self.left_frame)
        self.density_entry.insert(0, "7.85")
        self.density_entry.pack(fill="x", padx=20, pady=(0, 10))

        # 実行ボタン
        self.analyze_btn = ctk.CTkButton(self.left_frame, text="解析を実行", fg_color="#2FA572", hover_color="#107C41", command=self.start_analysis_thread)
        self.analyze_btn.pack(fill="x", padx=20, pady=20)

        # Excel出力ボタン (解析完了まで無効化)
        self.export_btn = ctk.CTkButton(self.left_frame, text="Excel / CSVに出力", fg_color="#0D6EFD", hover_color="#0b5ed7", command=self.export_to_excel, state="disabled")
        self.export_btn.pack(fill="x", padx=20, pady=(0, 20))

        # ステータス
        self.status_label = ctk.CTkLabel(self.left_frame, text="待機中...", font=ctk.CTkFont(slant="italic"))
        self.status_label.pack(side="bottom", padx=20, pady=20)

        # 右側パネル
        self.right_frame = ctk.CTkFrame(self, corner_radius=10)
        self.right_frame.grid(row=0, column=1, padx=(0, 20), pady=20, sticky="nsew")
        self.right_frame.grid_columnconfigure(0, weight=1)
        self.right_frame.grid_rowconfigure(1, weight=1)

        self.preview_label = ctk.CTkLabel(self.right_frame, text="図面プレビュー", fg_color="gray20", corner_radius=10, height=300)
        self.preview_label.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        # Treeview
        self.tree_frame = ctk.CTkFrame(self.right_frame)
        self.tree_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)

        cols = ("part_number", "part_name", "material", "dimensions", "volume", "weight", "notes")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show="headings")
        
        headings = {
            "part_number": {"text": "品番", "width": 50, "anchor": "w"},
            "part_name": {"text": "部品名/要素名", "width": 140, "anchor": "w"},
            "material": {"text": "材質 (比重)", "width": 100, "anchor": "w"},
            "dimensions": {"text": "読み取り寸法", "width": 140, "anchor": "w"},
            "volume": {"text": "体積 (mm³)", "width": 100, "anchor": "e"},
            "weight": {"text": "重量 (kg)", "width": 100, "anchor": "e"},
            "notes": {"text": "備考 (計算式等)", "width": 250, "anchor": "w"}
        }
        
        for k, info in headings.items():
            self.tree.heading(k, text=info["text"])
            self.tree.column(k, width=info["width"], anchor=info["anchor"])
        
        scroll = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        self.total_label = ctk.CTkLabel(self.right_frame, text="全体合計: --- mm³ / --- kg", font=ctk.CTkFont(size=18, weight="bold"))
        self.total_label.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="e")

    def select_file(self):
        path = filedialog.askopenfilename(title="図面ファイルを選択", filetypes=[("PDF/Image", "*.pdf *.png *.jpg *.jpeg *.bmp")])
        if path:
            self.selected_file_path = path
            self.file_path_display.configure(text=os.path.basename(path))

    def update_status(self, text):
        self.after(0, lambda: self.status_label.configure(text=text))

    def open_api_settings(self):
        plan = self.api_settings.get("plan", "free")
        self.api_settings[f"{plan}_key"] = self.api_key_entry.get().strip()
        APISettingsDialog(self, self.api_settings, self.on_api_settings_saved)

    def on_api_settings_saved(self, new_settings):
        self.api_settings = new_settings
        plan = new_settings["plan"]
        key = new_settings.get(f"{plan}_key", "")
        self.api_key_entry.delete(0, "end")
        self.api_key_entry.insert(0, key)
        self.plan_indicator.configure(text="🟢 無料枠 (Free)" if plan=="free" else "🔵 課金枠 (Paid)", 
                                     text_color="#198754" if plan=="free" else "#0D6EFD")
        self.save_settings()

    def update_preview(self, pil_image):
        temp_img = pil_image.copy()
        temp_img.thumbnail((600, 300))
        ctk_img = ctk.CTkImage(light_image=temp_img, dark_image=temp_img, size=temp_img.size)
        self.after(0, lambda: self.preview_label.configure(image=ctk_img, text=""))

    def start_analysis_thread(self):
        if not self.selected_file_path:
            messagebox.showwarning("警告", "ファイルを選択してください。", parent=self)
            return
        if not self.api_key_entry.get().strip():
            messagebox.showwarning("警告", "APIキーを入力してください。", parent=self)
            return

        self.analyze_btn.configure(state="disabled")
        self.export_btn.configure(state="disabled") # 解析中は無効化
        self.tree.delete(*self.tree.get_children())
        self.total_label.configure(text="全体合計: 解析中...")
        threading.Thread(target=self.run_analysis, daemon=True).start()

    def _parse_json_response(self, result_text):
        """モデルからのテキストレスポンスからJSONを抽出し、パースするヘルパー"""
        start = result_text.find('{')
        end = result_text.rfind('}')
        
        if start != -1:
            if end != -1 and end > start:
                json_str = result_text[start:end+1]
            else:
                json_str = result_text[start:]
        else:
            json_str = result_text

        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            self.update_status("JSONの自己修復を試みています...")
            return self._repair_and_parse_json(json_str)

    def run_analysis(self):
        try:
            plan = self.api_settings.get("plan", "free")
            prefix = f"{plan}_"
            api_key = self.api_key_entry.get().strip()
            density = self.density_entry.get().strip() or "7.85"
            
            # 各ステップのモデルを取得
            model_name_step1 = self.api_settings.get(f"{prefix}model_step1", "gemini-3-flash")
            model_name_step2 = self.api_settings.get(f"{prefix}model_step2", "gemini-3-flash")
            model_name_step3 = self.api_settings.get(f"{prefix}model_step3", "gemini-3.1-pro-preview")
            
            # モデル別パフォーマンス辞書を取得
            model_perf_dict = self.api_settings.get(f"{prefix}model_perf", {})
            
            # RPM/スレッドを取得するヘルパー
            def get_perf(m_id):
                perf = model_perf_dict.get(m_id, {})
                # 見つからなかった場合のフォールバックロジック
                is_pro = "pro" in m_id.lower()
                def_rpm = 2 if (is_pro and plan == "free") else (150 if is_pro else (15 if plan == "free" else 300))
                def_thr = 1 if (is_pro and plan == "free") else (5 if is_pro else (1 if plan == "free" else 5))
                return perf.get("rpm", def_rpm), perf.get("threads", def_thr)
                
            rpm_step1, threads_step1 = get_perf(model_name_step1)
            rpm_step2, threads_step2 = get_perf(model_name_step2)
            rpm_step3, threads_step3 = get_perf(model_name_step3)
            
            # 1. 画像とテキストデータの準備
            ext = os.path.splitext(self.selected_file_path)[1].lower()
            text_pdf = ""
            
            if ext == '.pdf':
                doc = fitz.open(self.selected_file_path)
                page = doc.load_page(0)
                pix = page.get_pixmap(dpi=400) # 高解像度でレンダリング
                raw_image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                
                # Visual Grounding 用にOpenCV形式へ変換
                annotated_img_cv = cv2.cvtColor(np.array(raw_image), cv2.COLOR_RGB2BGR)
                scale_factor = 400 / 72.0 # pdfplumberのデフォルト(72dpi)からのスケール変換
                
                self.update_status("ベクターデータを解析・グラウンディング画像生成中...")
                text_data_list = []
                with pdfplumber.open(self.selected_file_path) as pdf:
                    p0 = pdf.pages[0]
                    words = p0.extract_words()
                    
                    for idx, w in enumerate(words):
                        text = w['text']
                        # ID付きのテキストデータリストを作成
                        text_data_list.append(f"[ID:{idx}] X:{w['x0']:.1f}, Y:{w['top']:.1f} TEXT: {text}")
                        
                        # 画像にバウンディングボックスとIDを描画
                        x0 = int(w['x0'] * scale_factor)
                        y0 = int(w['top'] * scale_factor)
                        x1 = int(w['x1'] * scale_factor)
                        y1 = int(w['bottom'] * scale_factor)
                        
                        # 赤枠
                        cv2.rectangle(annotated_img_cv, (x0, y0), (x1, y1), (0, 0, 255), 2)
                        # 青文字ID
                        cv2.putText(annotated_img_cv, str(idx), (x0, max(0, y0 - 5)), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (255, 0, 0), 2)

                text_pdf = "\n".join(text_data_list)
                
                # アノテーション済みの画像をPillow形式に戻してAIへ渡す
                self.processed_image_pil = Image.fromarray(cv2.cvtColor(annotated_img_cv, cv2.COLOR_BGR2RGB))
            else:
                # 画像ファイルの場合のフォールバック
                raw_image = Image.open(self.selected_file_path).convert("RGB")
                self.processed_image_pil = raw_image

            self.update_preview(self.processed_image_pil)

            genai.configure(api_key=api_key)
            
            safety = None
            if self.api_settings.get(f"{prefix}safety", True):
                safety = [{"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                          {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                          {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                          {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}]

            tokens_val = self.api_settings.get(f"{prefix}tokens", 65535)
            if tokens_val > 65535:
                tokens_val = 65535

            common_generation_config_args = {
                "temperature": self.api_settings.get(f"{prefix}temp", 0.0),
                "max_output_tokens": tokens_val
            }
            
            custom_prompts = ' '.join(self.api_settings.get(f'{prefix}prompts', []))

            # ==========================================
            # ステップ1: 部品一覧の抽出
            # ==========================================
            self.update_status(f"ステップ1: 部品一覧を抽出中 ({model_name_step1})...")
            step1_prompt = f"""
            あなたは機械図面の解析エキスパートです。
            提供された図面画像とテキストデータから、「部品の一覧」のみをすべて抽出してください。
            ここでは各部品の寸法や要素の詳細は不要です。品番、品名、材質、比重のみを特定してください。
            
            【重要な前提条件（Visual Grounding）】
            - 画像には、テキストが抽出された正確な位置を示す「赤い枠」と「青い文字のID番号」が描画されています。
            - 以下のテキストデータ（[ID:〇〇] X座標, Y座標 TEXT: 文字列）は、画像上の青いID番号と完全にリンクしています。
            - 単純な文字の羅列ではなく、画像の「視覚的なレイアウト（表の構造や枠線）」と「テキストデータのID」を必ず照らし合わせて、正確に意味を読み取ってください。
            - デフォルト比重: {density} g/cm³ （材質記載がない場合に適用）
            
            テキストデータ:
            {text_pdf}
            
            ユーザーからの追加指示: {custom_prompts}
            """
            
            step1_config = genai.types.GenerationConfig(
                response_mime_type="application/json",
                response_schema=PartListResult,
                **common_generation_config_args
            )
            
            model_step1_instance = genai.GenerativeModel(model_name_step1)
            response1 = model_step1_instance.generate_content([step1_prompt, self.processed_image_pil], generation_config=step1_config, safety_settings=safety)
            if not response1.candidates: raise Exception("ステップ1: AIからの応答が空でした。")
            
            part_list_data = self._parse_json_response(response1.text)
            if not part_list_data: raise Exception("ステップ1: JSONの解析に失敗しました。")
            
            extracted_parts_basic = part_list_data.get('parts', [])
            final_parts = []

            # ==========================================
            # ステップ2: 個別部品の要素と寸法の抽出 (三面図解析 / マルチスレッド処理)
            # ==========================================
            model_step2_instance = genai.GenerativeModel(model_name_step2)
            
            def process_part(idx, p_basic):
                p_name = p_basic.get("part_name", "不明")
                p_num = p_basic.get("part_number", "")
                self.update_status(f"ステップ2: 部品詳細(三面図)を並列解析中... ({idx+1}/{len(extracted_parts_basic)}) {p_name}")
                
                step2_prompt = f"""
                あなたは機械図面の解析エキスパートです。
                提供された図面画像とテキストデータから、指定された特定の部品の「構成要素」と「三面図から導き出される3次元寸法」を漏れなく抽出し、体積計算式を作成してください。
                
                【対象部品】
                品番: {p_num}
                品名: {p_name}
                
                【作業手順 (三面図からの3D形状復元と線種の解析)】
                1. 対象部品を構成する要素（本体ベース、リブ、ボス、貫通穴など）に分解する。
                2. 図面の「線種」を正確に識別し、以下のように立体形状の解釈に反映させること。
                   - 実線（太線）: 見える外形線（部材の実際の外形）
                   - 破線: 隠れ線（内部の空洞、裏側の形状など。マイナス計算の強い根拠となる）
                   - 一点鎖線: 中心線、対称軸、ピッチ線（円柱や穴の中心を示す）
                   - 二点鎖線: 想像線、隣接部品、可動部の移動限界など（体積計算の対象外とする）
                3. 図面が「三面図（正面図、平面図、側面図など）」で描かれていることを前提とし、各投影図をまたいで寸法情報を統合する。
                   （例：正面図から「幅」と「高さ」を読み取り、平面図または側面図から「奥行き・厚み」を読み取って、頭の中で立体形状を復元する）
                4. 画像の「赤い枠と青いID番号」と、テキストデータのIDを照合し、他の部品の寸法を誤って拾わないように厳密に注意すること。
                5. 「calculation_formula」に、三面図から復元した寸法に基づく体積(mm³)を導き出す純粋な計算式を記述する。
                   ドリル穴、ザグリ、切り欠きなどの「空間（除去される部分）」は、必ずマイナス（引き算）として式を構築すること。
                   （例: 直方体ベースから円柱穴を引くなら "150 * 100 * 10 - (10/2)**2 * 3.14159 * 10"）
                
                テキストデータ:
                {text_pdf}
                
                ユーザーからの追加指示: {custom_prompts}
                """
                
                step2_config = genai.types.GenerationConfig(
                    response_mime_type="application/json",
                    response_schema=PartElementsResult,
                    **common_generation_config_args
                )
                
                # 自動リトライ機能（通信集中による429エラー等を防止）
                max_retries = 3
                response2 = None
                for attempt in range(max_retries):
                    try:
                        response2 = model_step2_instance.generate_content([step2_prompt, self.processed_image_pil], generation_config=step2_config, safety_settings=safety)
                        break
                    except Exception as e:
                        if attempt == max_retries - 1:
                            print(f"API Error on part {p_name}: {e}")
                            break
                        time.sleep((2 ** attempt) + 1) # 指数バックオフによる待機
                
                elements = []
                if response2:
                    elements_data = self._parse_json_response(response2.text)
                    elements = elements_data.get('elements', []) if elements_data else []
                
                complete_part = {
                    "part_number": p_basic.get("part_number", ""),
                    "part_name": p_basic.get("part_name", ""),
                    "material": p_basic.get("material", ""),
                    "density_g_cm3": p_basic.get("density_g_cm3", 0.0),
                    "elements": elements
                }
                return complete_part

            # 設定から取得したスレッド数を利用して並列実行
            max_workers = threads_step2 if threads_step2 > 0 else 1
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                # 順序を維持するために、submitした順番のリストを作成して順番に処理を待つ
                futures = [executor.submit(process_part, idx, p) for idx, p in enumerate(extracted_parts_basic)]
                for future in futures:
                    try:
                        result = future.result()
                        final_parts.append(result)
                    except Exception as e:
                        print(f"Thread execution error: {e}")

            # ==========================================
            # ステップ3: 最終検証 (抽出漏れの確認)
            # ==========================================
            self.update_status(f"ステップ3: 三面図からの抽出漏れがないか再検証中 ({model_name_step3})...")
            
            extracted_summary = "\n".join([f"- 品番:{p['part_number']} 品名:{p['part_name']}" for p in final_parts])
            
            step3_prompt = f"""
            あなたは機械図面の解析エキスパートです。最終チェックを行います。
            提供された図面画像およびテキストデータと、これまでに抽出した以下の部品リストを比較し、「まだ抽出されていない（見落としている）部品」が存在しないか厳密に確認してください。
            
            【すでに抽出済みの部品リスト】
            {extracted_summary}
            
            もし見落としている部品があれば、その部品情報（品番、品名、材質、比重）と、その構成要素（寸法、計算式等）を三面図解釈ルールに従って全て抽出して出力してください。
            画像の「青いID番号」とテキストを照らし合わせ、未処理の重要な情報ブロックが残っていないか確認してください。
            見落としが一つも無い場合（完璧な場合）は、「missing_parts」を空のリスト（[]）として返却してください。
            
            【前提条件・製図ルール】
            - デフォルト比重: {density} g/cm³ （材質記載がない場合に適用）
            - 線種の識別: 実線(外形)、破線(隠れ線/内部空間)、一点鎖線(中心軸)、二点鎖線(計算除外)を厳密に区別して解釈に反映すること。
            
            テキストデータ:
            {text_pdf}
            
            ユーザーからの追加指示: {custom_prompts}
            """
            
            step3_config = genai.types.GenerationConfig(
                response_mime_type="application/json",
                response_schema=MissingPartsResult,
                **common_generation_config_args
            )
            
            model_step3_instance = genai.GenerativeModel(model_name_step3)
            response3 = model_step3_instance.generate_content([step3_prompt, self.processed_image_pil], generation_config=step3_config, safety_settings=safety)
            verification_data = self._parse_json_response(response3.text)
            
            if verification_data and 'missing_parts' in verification_data:
                missing_parts = verification_data['missing_parts']
                if missing_parts:
                    self.update_status(f"ステップ3: {len(missing_parts)}件の抽出漏れを発見し、追加しました。")
                    final_parts.extend(missing_parts)
                else:
                    self.update_status("ステップ3: 見落としなしを確認しました。")

            # --- 最終結果の構築とUI表示 ---
            result_data = {"parts": final_parts}

            self.after(0, lambda: self.display_results(result_data))
            self.update_status("解析完了")

        except Exception as e:
            self.update_status("エラーが発生しました。")
            self.after(0, lambda e=e: self.total_label.configure(text="全体合計: --- mm³ / --- kg"))
            self.after(0, lambda e=e: messagebox.showerror("解析エラー", f"通信・パース中にエラーが発生しました。\n詳細:\n{str(e)}", parent=self))
        finally:
            self.after(0, lambda: self.analyze_btn.configure(state="normal"))

    def _repair_and_parse_json(self, json_str):
        """途切れたJSON文字列を強引に修復してパースを試みるベストエフォート機能"""
        for i in range(len(json_str), max(0, len(json_str) - 200), -1):
            temp_str = json_str[:i]
            
            stack = []
            in_string = False
            escape = False
            for char in temp_str:
                if escape:
                    escape = False
                    continue
                if char == '\\':
                    escape = True
                    continue
                if char == '"':
                    in_string = not in_string
                    continue
                if not in_string:
                    if char in '{[':
                        stack.append(char)
                    elif char == '}':
                        if stack and stack[-1] == '{':
                            stack.pop()
                    elif char == ']':
                        if stack and stack[-1] == '[':
                            stack.pop()
            
            repaired = temp_str
            if in_string:
                repaired += '"'
            
            while stack:
                char = stack.pop()
                if char == '{':
                    repaired += '}'
                elif char == '[':
                    repaired += ']'
                    
            try:
                parsed = json.loads(repaired)
                return parsed
            except json.JSONDecodeError:
                continue
                
        return None

    def _safe_float(self, val):
        """AIが数値を文字列（例: '15.0'）として返却した場合の型エラーを防ぐヘルパー"""
        if val is None or val == "":
            return 0.0
        try:
            if isinstance(val, str):
                val = val.replace(',', '')
            return float(val)
        except (ValueError, TypeError):
            return 0.0

    def _evaluate_formula(self, formula_str):
        """安全に数式文字列を評価して計算結果を返すヘルパー関数"""
        if not formula_str:
            return 0.0
        try:
            allowed_chars = set("0123456789+-*/(). ")
            sanitized = "".join(c for c in formula_str if c in allowed_chars)
            result = eval(sanitized)
            return float(result)
        except Exception as e:
            print(f"Formula evaluation error: {e} for formula: '{formula_str}'")
            return 0.0

    def display_results(self, data):
        """Treeviewに結果を表示し、Python側で体積・重量の計算を実行する"""
        total_overall_volume = 0.0
        total_overall_weight = 0.0
        
        for part in data.get('parts', []):
            p_num = str(part.get('part_number') or "")
            p_name = str(part.get('part_name') or "不明な部品")
            p_mat = str(part.get('material') or "")
            
            p_den = self._safe_float(part.get('density_g_cm3'))
            
            part_total_volume = 0.0
            for elem in part.get('elements', []):
                e_calc = str(elem.get('calculation_formula') or "")
                e_vol = self._evaluate_formula(e_calc)
                part_total_volume += e_vol
                elem['calculated_volume_mm3'] = e_vol
                elem['calculated_weight_kg'] = e_vol * p_den / 1000000

            part_total_weight = part_total_volume * p_den / 1000000
            
            total_overall_volume += part_total_volume
            total_overall_weight += part_total_weight
            
            if p_mat and p_den > 0.0:
                mat_display = f"{p_mat} ({p_den})"
            elif p_den > 0.0:
                mat_display = f"不明 ({p_den})"
            elif p_mat:
                mat_display = p_mat
            else:
                mat_display = "---"
                
            p_vol_str = f"{part_total_volume:,.1f}" if part_total_volume > 0.0 else "---"
            p_weight_str = f"{part_total_weight:,.2f}" if part_total_weight > 0.0 else "---"
            
            p_id = self.tree.insert("", "end", values=(
                p_num, 
                p_name, 
                mat_display,
                "", 
                p_vol_str, 
                p_weight_str, 
                ""
            ), open=True)
            
            for elem in part.get('elements', []):
                e_name = str(elem.get('element_name') or "要素")
                e_dim = str(elem.get('dimensions') or "")
                e_calc = str(elem.get('calculation_formula') or "")
                
                e_vol = elem.get('calculated_volume_mm3', 0.0)
                e_weight = elem.get('calculated_weight_kg', 0.0)
                e_notes = str(elem.get('notes') or "")
                
                combined_notes = e_notes
                if e_calc:
                    combined_notes = f"[式: {e_calc}] {e_notes}".strip()
                    
                e_vol_str = f"{e_vol:,.1f}" if e_vol > 0.0 else "---"
                e_weight_str = f"{e_weight:,.3f}" if e_weight > 0.0 else "---"
                
                self.tree.insert(p_id, "end", values=(
                    "", 
                    f"  ├ {e_name}", 
                    "",
                    e_dim,
                    e_vol_str,
                    e_weight_str,
                    combined_notes
                ))

        vol_str = f"{total_overall_volume:,.1f}" if total_overall_volume > 0.0 else "---"
        weight_str = f"{total_overall_weight:,.2f}" if total_overall_weight > 0.0 else "---"
        
        self.total_label.configure(text=f"全体合計: {vol_str} mm³ / {weight_str} kg")
        
        # データを保持してエクスポートボタンを有効化
        self.last_result_data = data
        self.export_btn.configure(state="normal")

    def export_to_excel(self):
        """保持している解析結果のデータをExcel(.xlsx)またはCSV形式で出力する"""
        if not self.last_result_data:
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")],
            title="保存先を選択"
        )
        if not file_path:
            return

        try:
            rows = []
            for part in self.last_result_data.get('parts', []):
                p_num = str(part.get('part_number') or "")
                p_name = str(part.get('part_name') or "不明な部品")
                p_mat = str(part.get('material') or "")
                p_den = self._safe_float(part.get('density_g_cm3'))
                
                for elem in part.get('elements', []):
                    e_name = str(elem.get('element_name') or "")
                    e_dim = str(elem.get('dimensions') or "")
                    e_calc = str(elem.get('calculation_formula') or "")
                    e_vol = elem.get('calculated_volume_mm3', 0.0)
                    e_weight = elem.get('calculated_weight_kg', 0.0)
                    e_notes = str(elem.get('notes') or "")
                    
                    combined_notes = e_notes
                    if e_calc:
                        combined_notes = f"[式: {e_calc}] {e_notes}".strip()
                        
                    rows.append({
                        "品番": p_num,
                        "部品名": p_name,
                        "材質": p_mat,
                        "比重 (g/cm³)": p_den,
                        "要素名": e_name,
                        "読み取り寸法": e_dim,
                        "体積 (mm³)": e_vol,
                        "重量 (kg)": e_weight,
                        "備考 (計算式等)": combined_notes
                    })

            if not rows:
                messagebox.showwarning("警告", "出力するデータがありません。", parent=self)
                return

            df = pd.DataFrame(rows)
            
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='解析結果')
                    
                    # 見やすくするための列幅の自動調整
                    worksheet = writer.sheets['解析結果']
                    for idx, col in enumerate(df.columns):
                        max_len = max(
                            df[col].astype(str).map(len).max(),
                            len(col)
                        )
                        worksheet.column_dimensions[chr(65 + idx)].width = min(max_len * 1.8, 60)
            
            messagebox.showinfo("出力完了", f"解析結果を保存しました。\n{file_path}", parent=self)
            
        except Exception as e:
            messagebox.showerror("出力エラー", f"ファイルの保存中にエラーが発生しました。\n詳細:\n{str(e)}", parent=self)

if __name__ == "__main__":
    app = DrawingAnalysisApp()
    app.mainloop()