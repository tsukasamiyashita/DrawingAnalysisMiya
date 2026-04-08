import customtkinter as ctk
from tkinter import messagebox
import threading
import time
import webbrowser
import google.generativeai as genai

# --- UI共通コンポーネント (チェックボックスリスト) ---
class CTkScrollableCheckboxList(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.items = []

    def set_items(self, item_texts):
        for item in self.items:
            item["cb"].destroy()
        self.items.clear()
        for text in item_texts:
            self.add_item(text)

    def add_item(self, text):
        var = ctk.StringVar(value="")
        cb = ctk.CTkCheckBox(self, text=text, variable=var, onvalue=text, offvalue="")
        cb.pack(anchor="w", padx=5, pady=2, fill="x")
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

# --- API設定ダイアログ ---
class APISettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_settings, on_save_callback):
        super().__init__(parent)
        self.title("⚙️ AI詳細設定 (Gemini API)")
        
        # 画面サイズ調整
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        dialog_w = min(1100, screen_w - 40)
        dialog_h = min(750, screen_h - 80)
        self.geometry(f"{dialog_w}x{dialog_h}")
        self.transient(parent)
        self.grab_set()

        self.settings = current_settings.copy()
        self.on_save_callback = on_save_callback
        
        # 1.5などの古いモデルを完全に排除し、最新モデルをデフォルトに
        default_models = [
            ("Gemini 3 Flash", "gemini-3-flash"),
            ("Gemini 3.1 Pro Preview", "gemini-3.1-pro-preview"),
            ("Gemini 3.1 Flash-Lite Preview", "gemini-3.1-flash-lite-preview"),
            ("Gemini 2.5 Flash", "gemini-2.5-flash"),
            ("Gemini 2.5 Pro", "gemini-2.5-pro")
        ]
        
        # 保存されているリストに古い1.5が含まれていればデフォルトで上書き
        saved_list = self.settings.get("models_list", [])
        if not saved_list or any("1.5" in item[1] for item in saved_list):
            self.models_list = default_models
        else:
            self.models_list = saved_list

        # 各モデルの推奨制限を保持する辞書
        self.model_limits = {}
        for m_display, m_id in self.models_list:
            is_pro = "pro" in m_id.lower()
            self.model_limits[m_id] = {
                "free_rpm": 2 if is_pro else 15,
                "free_threads": 1,
                "paid_rpm": 150 if is_pro else 300,
                "paid_threads": 5,
                "max_tokens": 8192
            }

        # 変数の初期化
        self.plan_var = ctk.StringVar(value=self.settings.get("plan", "free"))
        
        self.vars = {
            "free": {
                "key": ctk.StringVar(value=self.settings.get("free_key", "")),
                "model_step1": ctk.StringVar(value=self.settings.get("free_model_step1", "gemini-3-flash")),
                "model_step2": ctk.StringVar(value=self.settings.get("free_model_step2", "gemini-3-flash")),
                "model_step3": ctk.StringVar(value=self.settings.get("free_model_step3", "gemini-3.1-pro-preview")),
                "perf_vars": {}, # {model_id: {"rpm": IntVar, "threads": IntVar}} を格納
                "temp": ctk.DoubleVar(value=self.settings.get("free_temp", 0.0)),
                "tokens": ctk.IntVar(value=self.settings.get("free_tokens", 65535)),
                "safety": ctk.BooleanVar(value=self.settings.get("free_safety", True)),
                "prompts": self.settings.get("free_prompts", [])
            },
            "paid": {
                "key": ctk.StringVar(value=self.settings.get("paid_key", "")),
                "model_step1": ctk.StringVar(value=self.settings.get("paid_model_step1", "gemini-3-flash")),
                "model_step2": ctk.StringVar(value=self.settings.get("paid_model_step2", "gemini-3-flash")),
                "model_step3": ctk.StringVar(value=self.settings.get("paid_model_step3", "gemini-3.1-pro-preview")),
                "perf_vars": {},
                "temp": ctk.DoubleVar(value=self.settings.get("paid_temp", 0.0)),
                "tokens": ctk.IntVar(value=self.settings.get("paid_tokens", 65535)),
                "safety": ctk.BooleanVar(value=self.settings.get("paid_safety", True)),
                "prompts": self.settings.get("paid_prompts", [])
            }
        }
        
        self.saved_prompts = self.settings.get("saved_prompts", [])
        self.fav_lists = []
        self.tab_controls = [] # 各タブのコンボボックスと再描画関数を保持
        self.perf_render_funcs = [] # モデル更新時に各タブのスクロールUIを再描画するため

        self._setup_ui()

    def _setup_ui(self):
        # --- 下部ボタン領域 ---
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(side="bottom", fill="x", pady=(10, 15))
        
        ctk.CTkButton(btn_frame, text="設定を適用して閉じる", fg_color="#2FA572", hover_color="#107C41", 
                      command=self.save_and_close, width=200).pack(side="right", padx=20)
        ctk.CTkButton(btn_frame, text="キャンセル", fg_color="gray", hover_color="darkgray", 
                      command=self.destroy, width=150).pack(side="right", padx=10)

        # --- スクロール可能なメイン領域 ---
        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_frame.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        lbl_title = ctk.CTkLabel(self.scroll_frame, text="Gemini API 詳細設定", font=ctk.CTkFont(size=20, weight="bold"), text_color="#0D6EFD")
        lbl_title.pack(pady=(10, 5))

        # --- 実行プランの選択 ---
        plan_frame = ctk.CTkFrame(self.scroll_frame, corner_radius=8, border_width=1, border_color="gray70")
        plan_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(plan_frame, text=" 実行プランの選択 ", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(5, 0))
        plan_inner = ctk.CTkFrame(plan_frame, fg_color="transparent")
        plan_inner.pack(anchor="w", padx=10, pady=(5, 10))
        
        ctk.CTkLabel(plan_inner, text="実際に抽出で使用するプランを選んでください:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(plan_inner, text="無料枠 (Free Tier)", variable=self.plan_var, value="free").pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(plan_inner, text="課金枠 (Paid Tier)", variable=self.plan_var, value="paid").pack(side="left")

        # --- タブ ---
        self.tabview = ctk.CTkTabview(self.scroll_frame)
        self.tabview.pack(fill="both", expand=True, padx=15, pady=10)
        
        tab_free = self.tabview.add("🟢 無料枠 (Free Tier) の設定")
        tab_paid = self.tabview.add("🔵 課金枠 (Paid Tier) の設定")

        self.build_tab(tab_free, "free")
        self.build_tab(tab_paid, "paid")
        
        self.update_all_fav_lists()
        
        if self.plan_var.get() == "free":
            self.tabview.set("🟢 無料枠 (Free Tier) の設定")
        else:
            self.tabview.set("🔵 課金枠 (Paid Tier) の設定")

    def build_tab(self, parent_tab, plan_type):
        vars_dict = self.vars[plan_type]
        is_free = (plan_type == "free")

        # --- ① APIキー ---
        key_frame = ctk.CTkFrame(parent_tab, corner_radius=8, border_width=1, border_color="gray70")
        key_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(key_frame, text=" ① APIキー ", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(5, 0))
        
        key_inner = ctk.CTkFrame(key_frame, fg_color="transparent")
        key_inner.pack(fill="x", padx=10, pady=(5, 10))
        
        ctk.CTkLabel(key_inner, text=f"{'無料枠' if is_free else '課金枠'} 用のAPIキー:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 10))
        
        entry_key = ctk.CTkEntry(key_inner, textvariable=vars_dict["key"], width=400, show="*")
        entry_key.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        btn_toggle = ctk.CTkButton(key_inner, text="確認", width=60)
        btn_toggle.pack(side="left", padx=(0, 5))
        
        def toggle_key(e=entry_key, b=btn_toggle):
            if e.cget("show") == "*":
                e.configure(show="")
                b.configure(text="隠す")
            else:
                e.configure(show="*")
                b.configure(text="確認")
        btn_toggle.configure(command=toggle_key)

        def test_key(k_var, cb, btn):
            key = k_var.get().strip()
            if not key:
                messagebox.showwarning("警告", "APIキーが入力されていません。", parent=self)
                return
                
            current_display = cb.get()
            model_name = next((m[1] for m in self.models_list if m[0] == current_display), current_display)
            btn.configure(state="disabled", text="通信中...")

            def run_test():
                try:
                    genai.configure(api_key=key)
                    model = genai.GenerativeModel(model_name)
                    model.generate_content(f"Connection Test: {time.time()}")
                    self.after(0, lambda: messagebox.showinfo("テスト成功", f"APIキーは正しく認識されました。\nモデル「{model_name}」による通信は正常です！", parent=self))
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("通信エラー", f"通信に問題が発生しました。\n詳細:\n{e}", parent=self))
                finally:
                    self.after(0, lambda: btn.configure(state="normal", text="テスト"))

            threading.Thread(target=run_test, daemon=True).start()

        # --- 中間コンテナ上部（左右分割） ---
        middle_top_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        middle_top_frame.pack(fill="x", padx=10, pady=5)
        middle_top_frame.grid_columnconfigure(0, weight=3)
        middle_top_frame.grid_columnconfigure(1, weight=2)

        # --- ② モデル設定 (左側) ---
        model_frame = ctk.CTkFrame(middle_top_frame, corner_radius=8, border_width=1, border_color="gray70")
        model_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        ctk.CTkLabel(model_frame, text=" ② 使用モデル設定 (ステップ別) ", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(5, 0))
        
        model_inner = ctk.CTkFrame(model_frame, fg_color="transparent")
        model_inner.pack(fill="x", padx=10, pady=(5, 2))
        
        current_plan_combos = []
        step_labels = ["Step1 (一覧抽出):", "Step2 (要素抽出):", "Step3 (最終検証):"]
        var_keys = ["model_step1", "model_step2", "model_step3"]
        display_names = [m[0] for m in self.models_list]
        
        for i, (label_text, v_key) in enumerate(zip(step_labels, var_keys)):
            step_frame = ctk.CTkFrame(model_inner, fg_color="transparent")
            step_frame.pack(fill="x", pady=2)
            
            ctk.CTkLabel(step_frame, text=label_text, font=ctk.CTkFont(weight="bold"), width=110, anchor="w").pack(side="left")
            
            cb = ctk.CTkComboBox(step_frame, values=display_names)
            cb.pack(side="left", fill="x", expand=True, padx=(0, 5))
            
            btn_test = ctk.CTkButton(step_frame, text="テスト", width=40)
            btn_test.configure(command=lambda k=vars_dict["key"], c=cb, b=btn_test: test_key(k, c, b))
            btn_test.pack(side="left")
            
            curr_id = vars_dict[v_key].get()
            matched_display = next((m[0] for m in self.models_list if m[1] == curr_id), curr_id)
            cb.set(matched_display)
            
            current_plan_combos.append(cb)
            
        action_inner = ctk.CTkFrame(model_frame, fg_color="transparent")
        action_inner.pack(fill="x", padx=10, pady=(10, 5))
        
        def fetch_models(k_var=vars_dict["key"], btn=None):
            key = k_var.get().strip()
            if not key:
                messagebox.showwarning("警告", "最新モデルを取得するには、APIキーを入力してください。", parent=self)
                return
            btn.configure(text="取得中...", state="disabled")
            
            def do_fetch():
                try:
                    genai.configure(api_key=key)
                    new_models = []
                    max_out_tokens = 8192
                    
                    for m in genai.list_models():
                        if "generateContent" in getattr(m, "supported_generation_methods", []) and "gemini" in m.name.lower():
                            name = m.name.replace("models/", "")
                            if not any(k in name.lower() for k in ["tts", "audio", "image", "vision", "embedding"]):
                                new_models.append((f"{m.display_name} ({name})", name))
                                
                                # トークンと制限を調べる
                                out_tokens = getattr(m, "output_token_limit", 8192)
                                if out_tokens > max_out_tokens:
                                    max_out_tokens = out_tokens
                                    
                                is_pro = "pro" in name.lower()
                                self.model_limits[name] = {
                                    "free_rpm": 2 if is_pro else 15,
                                    "free_threads": 1,
                                    "paid_rpm": 150 if is_pro else 300,
                                    "paid_threads": 5,
                                    "max_tokens": out_tokens
                                }
                            
                    if new_models:
                        self.models_list.clear()
                        self.models_list.extend(new_models)
                        
                        # 安全なトークン上限 (最大65535) にクリップ
                        safe_tokens = min(max_out_tokens, 65535)
                        
                        self.after(0, lambda: update_ui_after_fetch(safe_tokens))
                        self.after(0, lambda: messagebox.showinfo("更新完了", f"最新のモデルリスト ({len(self.models_list)}件) とAPI制限を取得し適用しました！\n選択肢と最大トークンが更新されました。", parent=self))
                    else:
                        self.after(0, lambda: messagebox.showinfo("情報", "取得可能なモデルが見つかりませんでした。", parent=self))
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("エラー", f"モデルリストの取得に失敗しました。\n詳細: {e}", parent=self))
                finally:
                    self.after(0, lambda: btn.configure(text="🌐 モデルリスト更新", state="normal"))

            def update_ui_after_fetch(new_tokens):
                display_names_updated = [m[0] for m in self.models_list]
                for controls in self.tab_controls:
                    cbs = controls["model_combos"]
                    for cb in cbs:
                        current_val = cb.get()
                        cb.configure(values=display_names_updated)
                        # 可能であれば元の選択を維持
                        cb.set(current_val)
                
                # 取得した最大トークンを無料枠・課金枠の両方に適用
                self.vars["free"]["tokens"].set(new_tokens)
                self.vars["paid"]["tokens"].set(new_tokens)
                
                # 両方のタブのパフォーマンス設定表示を再描画
                for render_func in self.perf_render_funcs:
                    render_func()

            threading.Thread(target=do_fetch, daemon=True).start()

        btn_fetch = ctk.CTkButton(action_inner, text="🌐 モデルリスト更新", width=140)
        btn_fetch.configure(command=lambda k=vars_dict["key"], b=btn_fetch: fetch_models(k, b))
        btn_fetch.pack(side="left")
        
        lbl_link = ctk.CTkLabel(action_inner, text="🔗 各モデルの特徴 (公式)", text_color="#0D6EFD", cursor="hand2", font=ctk.CTkFont(underline=True))
        lbl_link.pack(side="right")
        lbl_link.bind("<Button-1>", lambda e: webbrowser.open_new("https://ai.google.dev/gemini-api/docs/models/gemini"))


        # --- ③ AI抽出パラメータ設定 (右側) ---
        param_frame = ctk.CTkFrame(middle_top_frame, corner_radius=8, border_width=1, border_color="gray70")
        param_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        ctk.CTkLabel(param_frame, text=" ③ AI抽出パラメータ設定 ", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(5, 0))
        
        param_row1 = ctk.CTkFrame(param_frame, fg_color="transparent")
        param_row1.pack(fill="x", padx=10, pady=(15, 5))
        ctk.CTkLabel(param_row1, text="Temp:", font=ctk.CTkFont(weight="bold")).pack(side="left")
        ctk.CTkEntry(param_row1, textvariable=vars_dict["temp"], width=50).pack(side="left", padx=(5, 15))
        ctk.CTkLabel(param_row1, text="最大トークン:", font=ctk.CTkFont(weight="bold")).pack(side="left")
        ctk.CTkEntry(param_row1, textvariable=vars_dict["tokens"], width=80).pack(side="left", padx=(5, 0))

        param_row2 = ctk.CTkFrame(param_frame, fg_color="transparent")
        param_row2.pack(fill="x", padx=10, pady=5)
        ctk.CTkCheckBox(param_row2, text="安全フィルタ無効化", variable=vars_dict["safety"]).pack(side="left")

        def reset_param(t_var=vars_dict["temp"], tok_var=vars_dict["tokens"], s_var=vars_dict["safety"]):
            t_var.set(0.0)
            tok_var.set(65535)
            s_var.set(True)
            
        btn_reset_param = ctk.CTkButton(param_row2, text="🔄 推奨値", width=80, 
                                        command=lambda t=vars_dict["temp"], tok=vars_dict["tokens"], s=vars_dict["safety"]: reset_param(t, tok, s))
        btn_reset_param.pack(side="right", pady=(15, 0))


        # --- ④ パフォーマンス設定（横幅いっぱい） ---
        perf_container = ctk.CTkFrame(parent_tab, corner_radius=8, border_width=1, border_color="gray70")
        perf_container.pack(fill="x", padx=10, pady=(5, 5))
        ctk.CTkLabel(perf_container, text=" ④ パフォーマンス設定 (モデル別 API制限) ", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(5, 0))

        # スクロールエリアを廃止し、すべて展開表示される通常のFrameに変更
        perf_list_frame = ctk.CTkFrame(perf_container, fg_color="transparent")
        perf_list_frame.pack(fill="x", padx=10, pady=5)
        
        vars_dict["perf_vars"] = {} # {model_id: {"rpm": IntVar, "threads": IntVar}}
        
        def render_perf_settings():
            for widget in perf_list_frame.winfo_children():
                widget.destroy()
                
            model_perf_data = self.settings.get(f"{plan_type}_model_perf", {})
            
            for m_display, m_id in self.models_list:
                row_frame = ctk.CTkFrame(perf_list_frame, fg_color="transparent")
                row_frame.pack(fill="x", pady=2)
                
                # ラベル幅を固定して整列
                ctk.CTkLabel(row_frame, text=m_display, anchor="w", width=250).pack(side="left", padx=(0, 10))
                
                # IntVarの初期化 (既存の値があれば引き継ぐ)
                if m_id not in vars_dict["perf_vars"]:
                    saved_perf = model_perf_data.get(m_id)
                    if saved_perf:
                        def_rpm = saved_perf.get("rpm")
                        def_thr = saved_perf.get("threads")
                    else:
                        # 内部辞書（API制限）からデフォルトを適用
                        limits = self.model_limits.get(m_id, {})
                        def_rpm = limits.get("free_rpm" if is_free else "paid_rpm", 2 if "pro" in m_id.lower() and is_free else (150 if "pro" in m_id.lower() else (15 if is_free else 300)))
                        def_thr = limits.get("free_threads" if is_free else "paid_threads", 1 if "pro" in m_id.lower() and is_free else (5 if "pro" in m_id.lower() else (1 if is_free else 5)))
                    
                    vars_dict["perf_vars"][m_id] = {
                        "rpm": ctk.IntVar(value=def_rpm),
                        "threads": ctk.IntVar(value=def_thr)
                    }
                
                rpm_var = vars_dict["perf_vars"][m_id]["rpm"]
                thr_var = vars_dict["perf_vars"][m_id]["threads"]
                
                ctk.CTkLabel(row_frame, text="RPM:", font=ctk.CTkFont(size=11)).pack(side="left")
                ctk.CTkEntry(row_frame, textvariable=rpm_var, width=50).pack(side="left", padx=(2, 20))
                
                ctk.CTkLabel(row_frame, text="スレッド:", font=ctk.CTkFont(size=11)).pack(side="left")
                ctk.CTkEntry(row_frame, textvariable=thr_var, width=50).pack(side="left", padx=(2, 0))

        render_perf_settings()
        self.perf_render_funcs.append(render_perf_settings)

        perf_action_inner = ctk.CTkFrame(perf_container, fg_color="transparent")
        perf_action_inner.pack(fill="x", padx=10, pady=(5, 10))
        
        btn_show_limit = ctk.CTkButton(perf_action_inner, text="ℹ️ 制限と仕様を確認", fg_color="gray", hover_color="darkgray", 
                                       command=lambda f=is_free: self.show_limit_info(f))
        btn_show_limit.pack(side="left")

        def reset_perf():
            for m_id, p_vars in vars_dict["perf_vars"].items():
                limits = self.model_limits.get(m_id, {})
                def_rpm = limits.get("free_rpm" if is_free else "paid_rpm", 2 if "pro" in m_id.lower() and is_free else (150 if "pro" in m_id.lower() else (15 if is_free else 300)))
                def_thr = limits.get("free_threads" if is_free else "paid_threads", 1 if "pro" in m_id.lower() and is_free else (5 if "pro" in m_id.lower() else (1 if is_free else 5)))
                
                p_vars["rpm"].set(def_rpm)
                p_vars["threads"].set(def_thr)
                    
        btn_reset_perf = ctk.CTkButton(perf_action_inner, text="🔄 パフォーマンスの推奨値セット", width=80, 
                                       command=reset_perf)
        btn_reset_perf.pack(side="right")
        
        self.tab_controls.append({"model_combos": current_plan_combos})


        # --- ⑤ カスタムプロンプト ---
        prompt_frame = ctk.CTkFrame(parent_tab, corner_radius=8, border_width=1, border_color="gray70")
        prompt_frame.pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkLabel(prompt_frame, text=" ⑤ 独自の追加指示 (カスタムプロンプト) ", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(5, 0))

        input_inner = ctk.CTkFrame(prompt_frame, fg_color="transparent")
        input_inner.pack(fill="x", padx=10, pady=(5, 10))
        
        entry_new_prompt = ctk.CTkEntry(input_inner, placeholder_text="新しい指示を入力...")
        entry_new_prompt.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        def add_current_prompt(e=None):
            text = entry_new_prompt.get().strip()
            if text:
                current_list.add_item(text)
                entry_new_prompt.delete(0, "end")
                sync_current_to_var()

        entry_new_prompt.bind("<Return>", add_current_prompt)
        ctk.CTkButton(input_inner, text="＋ 追加", command=add_current_prompt, width=100).pack(side="left")

        lists_frame = ctk.CTkFrame(prompt_frame, fg_color="transparent")
        lists_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        lists_frame.grid_columnconfigure(0, weight=1)
        lists_frame.grid_columnconfigure(1, weight=1)

        left_frame = ctk.CTkFrame(lists_frame, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        current_list = CTkScrollableCheckboxList(left_frame, height=120, border_width=1)
        current_list.pack(fill="both", expand=True, pady=5)
        
        def sync_current_to_var():
            vars_dict["prompts"] = current_list.get_all_items()

        right_frame = ctk.CTkFrame(lists_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        fav_list = CTkScrollableCheckboxList(right_frame, height=120, border_width=1)
        fav_list.pack(fill="both", expand=True, pady=5)

        def add_fav_to_current():
            for text in fav_list.get_selected_items():
                current_list.add_item(text)
            sync_current_to_var()

        ctk.CTkButton(left_frame, text="🗑 削除", command=lambda: [current_list.remove_selected(), sync_current_to_var()], width=80).pack(side="left")
        ctk.CTkButton(right_frame, text="◀ 追加", command=add_fav_to_current, width=80).pack(side="left")

        current_list.set_items(vars_dict["prompts"])

    def update_all_fav_lists(self):
        for f_list in self.fav_lists:
            f_list.set_items(self.saved_prompts)

    # ℹ️ 制限と仕様を確認用のダイアログ表示
    def show_limit_info(self, is_f):
        info_win = ctk.CTkToplevel(self)
        info_win.title("Gemini API 制限と仕様一覧")
        info_win.geometry("950x700")
        info_win.transient(self)
        info_win.grab_set()
        
        scroll = ctk.CTkScrollableFrame(info_win)
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        ctk.CTkLabel(scroll, text="Gemini API 仕様・制限一覧", font=ctk.CTkFont(size=18, weight="bold"), text_color="#0D6EFD").pack(pady=(10, 20))
        
        def create_table(parent, headers, data, col_weights):
            frame = ctk.CTkFrame(parent, border_width=1, border_color="gray50")
            frame.pack(fill="x", padx=10, pady=(0, 20))
            
            for i, weight in enumerate(col_weights):
                frame.grid_columnconfigure(i, weight=weight)
                
            for col_idx, text in enumerate(headers):
                lbl = ctk.CTkLabel(frame, text=text, font=ctk.CTkFont(weight="bold"), fg_color="gray80", text_color="black", corner_radius=0, padx=5, pady=5)
                lbl.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
                
            for row_idx, row_data in enumerate(data, 1):
                for col_idx, text in enumerate(row_data):
                    lbl = ctk.CTkLabel(frame, text=text, fg_color="white", text_color="black", corner_radius=0, justify="left", anchor="nw", padx=5, pady=5)
                    lbl.grid(row=row_idx, column=col_idx, sticky="nsew", padx=1, pady=1)

        # テーブル1: プラン比較
        ctk.CTkLabel(scroll, text="▼ プラン比較", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=10)
        headers_plan = ["比較項目", "無料枠 (Free Tier)", "課金枠 (Paid Tier)"]
        data_plan = [
            ["利用料金", "完全無料（クレジットカード登録不要）", "従量課金（トークンと呼ばれるデータ量に応じて支払い）"],
            ["利用できるモデル", "3 Flash, 3.1 Flash-Lite, 2.5 Pro など", "すべてのモデルが利用可"],
            ["データの\nプライバシー", "入力データがGoogleのAI学習に利用される可能性がある", "入力データはAI学習に利用されない"]
        ]
        create_table(scroll, headers_plan, data_plan, [1, 3, 3])

        KNOWN_MODEL_INFO = {
            "gemini-3.1-pro-preview": {
                "limit": ["非常に厳しい (2 RPM未満など)\n[推奨: 1 RPM / 直列(1)]", "時期・モデルにより変動\n[推奨: 150 RPM / 並列(5)]"],
                "desc": ["最新鋭・最高精度モデル", "複雑な表の構造解析、かすれた手書き文字の正確な読み取り、論理推論", "複雑なレイアウトの図面、絶対にミスが許されないデータ抽出"]
            },
            "gemini-3-flash": {
                "limit": ["15 RPM, 1500 RPD\n[推奨: 12 RPM / 直列(1)]", "1000 RPM\n[推奨: 300 RPM / 並列(5)]"],
                "desc": ["高速・高性能バランス型", "スピードと精度の高い両立、画像認識（マルチモーダル）", "一般的な図面解析やPDFのテキスト・表抽出（デフォルト推奨）"]
            },
            "gemini-3.1-flash-lite-preview": {
                "limit": ["15 RPM, 1500 RPD\n[推奨: 12 RPM / 直列(1)]", "1000 RPM\n[推奨: 300 RPM / 並列(5)]"],
                "desc": ["最軽量・低コストモデル", "圧倒的な処理スピードと低コスト（Proの約1/8の価格）", "画質が良いPDFの単純なテキスト抽出、大量データを安価に処理したい場合"]
            },
            "gemini-2.5-flash": {
                "limit": ["15 RPM", "1000 RPM"],
                "desc": ["前世代の標準モデル", "過去の互換性維持のため", "-"]
            },
            "gemini-2.5-pro": {
                "limit": ["2 RPM", "360 RPM"],
                "desc": ["前世代の高精度モデル", "過去の互換性維持のため", "-"]
            }
        }

        data_limit = []
        data_model = []

        for m_display, m_id in self.models_list:
            info = KNOWN_MODEL_INFO.get(m_id)
            if not info:
                for known_id, known_data in KNOWN_MODEL_INFO.items():
                    if known_id in m_id:
                        info = known_data
                        break
            
            if info:
                data_limit.append([m_display, info["limit"][0], info["limit"][1]])
                data_model.append([m_display, info["desc"][0], info["desc"][1], info["desc"][2]])
            else:
                data_limit.append([m_display, "詳細は公式ドキュメントを参照", "詳細は公式ドキュメントを参照"])
                data_model.append([m_display, "APIから取得した追加モデル", "-", "最新機能を試したい場合"])

        # テーブル2: 各モデルの制限目安
        ctk.CTkLabel(scroll, text="▼ 各モデルの制限目安 (RPM と 推奨スレッド数)", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=10)
        headers_limit = ["モデル名", "無料枠の制限目安\n(RPM / スレッド数)", "課金枠の制限目安\n(RPM / スレッド数)"]
        create_table(scroll, headers_limit, data_limit, [2, 3, 3])

        # テーブル3: 特徴と適した用途
        ctk.CTkLabel(scroll, text="▼ 各モデルの特徴と適した用途", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=10)
        headers_model = ["モデル名", "特徴", "得意なこと", "適した用途"]
        create_table(scroll, headers_model, data_model, [2, 2, 3, 3])

        current_plan = "無料枠 (Free Tier)" if is_f else "課金枠 (Paid Tier)"
        ctk.CTkLabel(scroll, text=f"【現在表示中のプラン情報】\n{current_plan}", font=ctk.CTkFont(weight="bold"), text_color="#0D6EFD", 
                     fg_color="gray80", corner_radius=8, padx=10, pady=10).pack(fill="x", padx=10, pady=15)
                     
        ctk.CTkButton(scroll, text="閉じる", command=info_win.destroy, width=150, fg_color="gray").pack(pady=(0, 20))


    def save_and_close(self):
        new_settings = {
            "plan": self.plan_var.get(),
            "models_list": self.models_list,
            "saved_prompts": self.saved_prompts,
        }
        
        # 保存時に現在の情報を取得
        for plan_type, controls in zip(["free", "paid"], self.tab_controls):
            v = self.vars[plan_type]
            cbs = controls["model_combos"]
            
            # Step1~3 のモデル割り当て
            for step_idx, step_name in enumerate(["step1", "step2", "step3"]):
                current_display = cbs[step_idx].get()
                matched_id = next((m[1] for m in self.models_list if m[0] == current_display), current_display)
                new_settings[f"{plan_type}_model_{step_name}"] = matched_id
                
            # モデルごとのパフォーマンス (RPM/Threads) を辞書化して保存
            model_perf_dict = {}
            for m_id, p_vars in v["perf_vars"].items():
                model_perf_dict[m_id] = {
                    "rpm": p_vars["rpm"].get(),
                    "threads": p_vars["threads"].get()
                }
            new_settings[f"{plan_type}_model_perf"] = model_perf_dict
            
            new_settings[f"{plan_type}_key"] = v["key"].get().strip()
            new_settings[f"{plan_type}_temp"] = v["temp"].get()
            new_settings[f"{plan_type}_tokens"] = v["tokens"].get()
            new_settings[f"{plan_type}_safety"] = v["safety"].get()
            new_settings[f"{plan_type}_prompts"] = v["prompts"]

        self.on_save_callback(new_settings)
        self.destroy()