# -*- coding: utf-8 -*-
"""
DrawingAnalysisMiya モダン・ヘルプ＆ガイドセンター
常時ライトモード固定
"""

import os
import sys
import tkinter as tk
import tkinter.scrolledtext as st
import webbrowser

import customtkinter as ctk

from api_client import APP_NAME, APP_VERSION
from settings_dialog import (
    AZURE_HEADER_BG,
    BG_COLOR,
    BORDER_COLOR,
    CARD_BG,
    COLOR_DANGER,
    COLOR_SUCCESS,
    COLOR_WARNING,
    HEADER_BG,
    MUTED_TEXT,
    PRIMARY,
    PRIMARY_HOVER,
    TEXT_COLOR,
    center_window,
)

COLOR_PURPLE = "#6F42C1"
COLOR_INFO = "#0BACCE"

TAB_GUIDE = "🚀 使い方ガイド"
TAB_AI = "🤖 AI・API の準備"
TAB_TIPS = "📐 精度を上げるコツ"
TAB_FAQ = "❓ よくある質問 (FAQ)"
TAB_ABOUT = "📄 Readme / 情報"


def resource_path(relative_path):
    """PyInstaller (--onefile) 実行環境と通常実行環境のどちらでもリソースパスを解決する"""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)


def _load_readme_text():
    for fname in ("readme.md", "README.md"):
        candidates = [
            resource_path(fname),
            os.path.join(os.path.abspath(os.path.dirname(__file__)), fname),
            os.path.join(os.path.abspath("."), fname),
        ]
        for path in candidates:
            if os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        return f.read()
                except Exception:
                    pass
    return "Readme ファイルが見つかりませんでした。"


class HelpDialog(ctk.CTkToplevel):
    """
    ビジュアルで見やすい統合ヘルプ＆ガイドセンター（ライトモード固定）。
    initial_tab: 'guide' | 'ai' | 'tips' | 'faq' | 'about'
    """

    def __init__(self, parent, on_open_api_settings=None, settings_dir="", initial_tab="guide"):
        super().__init__(parent)
        self.title(f"📖 {APP_NAME} ヘルプ ＆ ガイドセンター")
        self.configure(fg_color=BG_COLOR)

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        win_w = min(1020, screen_w - 40)
        win_h = min(760, screen_h - 80)
        center_window(self, parent, win_w, win_h)
        self.minsize(820, 580)
        self.transient(parent)
        self.grab_set()

        self._on_open_api_settings = on_open_api_settings
        self._settings_dir = settings_dir
        self._win_w = win_w

        self._setup_ui(initial_tab)
        self.after(60, self._bring_to_front)

    def _bring_to_front(self):
        try:
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def _open_api_settings_and_close(self):
        self.destroy()
        if self._on_open_api_settings:
            self._on_open_api_settings()

    def _open_settings_folder(self):
        if not self._settings_dir:
            return
        try:
            os.makedirs(self._settings_dir, exist_ok=True)
            os.startfile(self._settings_dir)
        except Exception:
            pass

    # ==============================
    # 共通カード / テーブルビルダー
    # ==============================
    def _create_accent_card(
        self, parent, accent_color, badge_text, title_text, body_text, extra_builder=None
    ):
        """左側にカラーアクセントバーとバッジを備えたモダンな解説カードを作成する"""
        card = ctk.CTkFrame(
            parent,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        card.pack(fill="x", padx=8, pady=6)
        card.grid_columnconfigure(1, weight=1)

        accent_bar = ctk.CTkFrame(card, fg_color=accent_color, width=6, corner_radius=3)
        accent_bar.grid(row=0, column=0, sticky="ns", padx=(4, 0), pady=4)

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.grid(row=0, column=1, sticky="nsew", padx=14, pady=10)

        top_row = ctk.CTkFrame(body, fg_color="transparent")
        top_row.pack(fill="x", pady=(0, 6))

        if badge_text:
            badge_fg = "#212529" if accent_color == COLOR_WARNING else "white"
            badge = ctk.CTkLabel(
                top_row,
                text=f" {badge_text} ",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color=accent_color,
                text_color=badge_fg,
                corner_radius=5,
                padx=8,
                pady=2,
            )
            badge.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            top_row,
            text=title_text,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=TEXT_COLOR,
            anchor="w",
        ).pack(side="left", fill="x", expand=True)

        if body_text:
            ctk.CTkLabel(
                body,
                text=body_text,
                font=ctk.CTkFont(size=12),
                text_color=TEXT_COLOR,
                justify="left",
                anchor="w",
                wraplength=max(560, self._win_w - 160),
            ).pack(fill="x", pady=(0, 4))

        if extra_builder:
            extra_frame = ctk.CTkFrame(body, fg_color="transparent")
            extra_frame.pack(fill="x", pady=(6, 0))
            extra_builder(extra_frame)

    def _create_styled_table(
        self, parent, headers, rows, col_weights, header_bg=AZURE_HEADER_BG, header_fg=PRIMARY
    ):
        """見やすいストライプ付き比較テーブルを作成する"""
        table = ctk.CTkFrame(parent, fg_color=BORDER_COLOR, corner_radius=6)
        table.pack(fill="x", padx=8, pady=(4, 12))

        total_weight = sum(col_weights)
        usable_w = max(620, self._win_w - 120)

        for idx, w in enumerate(col_weights):
            table.grid_columnconfigure(idx, weight=w)

        for col_idx, h_text in enumerate(headers):
            col_wrap = max(90, int(usable_w * (col_weights[col_idx] / total_weight)) - 22)
            ctk.CTkLabel(
                table,
                text=h_text,
                font=ctk.CTkFont(size=12, weight="bold"),
                fg_color=header_bg,
                text_color=header_fg,
                corner_radius=0,
                padx=10,
                pady=8,
                justify="left",
                anchor="w",
                wraplength=col_wrap,
            ).grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)

        for r_idx, r_data in enumerate(rows, 1):
            row_bg = CARD_BG if r_idx % 2 == 1 else "#F8FAFC"
            for col_idx, val in enumerate(r_data):
                col_wrap = max(90, int(usable_w * (col_weights[col_idx] / total_weight)) - 22)
                ctk.CTkLabel(
                    table,
                    text=val,
                    font=ctk.CTkFont(size=11, weight="bold" if col_idx == 0 else "normal"),
                    fg_color=row_bg,
                    text_color=TEXT_COLOR,
                    corner_radius=0,
                    padx=10,
                    pady=8,
                    justify="left",
                    anchor="nw",
                    wraplength=col_wrap,
                ).grid(row=r_idx, column=col_idx, sticky="nsew", padx=1, pady=1)

    def _section_heading(self, parent, text, pady=(12, 4)):
        ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=PRIMARY,
            anchor="w",
        ).pack(fill="x", padx=8, pady=pady)

    # ==============================
    # UI 全体構築
    # ==============================
    def _setup_ui(self, initial_tab):
        # --- 1. 上部ヘッダーバナー ---
        header = ctk.CTkFrame(
            self,
            fg_color=CARD_BG,
            corner_radius=0,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        header.pack(side="top", fill="x")

        header_inner = ctk.CTkFrame(header, fg_color="transparent")
        header_inner.pack(fill="x", padx=18, pady=12)

        left_box = ctk.CTkFrame(header_inner, fg_color="transparent")
        left_box.pack(side="left", fill="x", expand=True)

        title_row = ctk.CTkFrame(left_box, fg_color="transparent")
        title_row.pack(anchor="w")

        ctk.CTkLabel(
            title_row,
            text=f"📖 {APP_NAME} ヘルプ ＆ ガイドセンター",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=PRIMARY,
        ).pack(side="left")

        ctk.CTkLabel(
            title_row,
            text=f" {APP_VERSION} ",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=AZURE_HEADER_BG,
            text_color=PRIMARY,
            corner_radius=5,
            padx=6,
            pady=2,
        ).pack(side="left", padx=(10, 0))

        ctk.CTkLabel(
            left_box,
            text="2D機械図面の重量・概算材料費・搬送トラック選定、切断寸法入り部材表、購入品・機械加工品抽出の手順を確認できます。",
            font=ctk.CTkFont(size=12),
            text_color=MUTED_TEXT,
        ).pack(anchor="w", pady=(4, 0))

        if self._on_open_api_settings:
            ctk.CTkButton(
                header_inner,
                text="⚙️ AI詳細設定を開く",
                fg_color=PRIMARY,
                hover_color=PRIMARY_HOVER,
                text_color="white",
                font=ctk.CTkFont(weight="bold"),
                width=165,
                command=self._open_api_settings_and_close,
            ).pack(side="right", padx=(12, 0))

        # --- 2. 下部フッター ---
        footer = ctk.CTkFrame(
            self,
            fg_color=CARD_BG,
            corner_radius=0,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        footer.pack(side="bottom", fill="x")

        footer_inner = ctk.CTkFrame(footer, fg_color="transparent")
        footer_inner.pack(fill="x", padx=18, pady=10)

        ctk.CTkLabel(
            footer_inner,
            text="💡 ヒント: 各タブ内はマウスホイールで上下にスクロールできます。",
            font=ctk.CTkFont(size=11),
            text_color=MUTED_TEXT,
        ).pack(side="left")

        ctk.CTkButton(
            footer_inner,
            text="閉じる",
            command=self.destroy,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            width=120,
        ).pack(side="right")

        # --- 3. タブビュー ---
        self.tabview = ctk.CTkTabview(
            self,
            fg_color=BG_COLOR,
            segmented_button_fg_color="#DEE2E6",
            segmented_button_selected_color=PRIMARY,
            segmented_button_selected_hover_color=PRIMARY_HOVER,
            segmented_button_unselected_color="#6C757D",
            segmented_button_unselected_hover_color="#5C636A",
            text_color="white",
            border_width=1,
            border_color=BORDER_COLOR,
        )
        self.tabview.pack(side="top", fill="both", expand=True, padx=12, pady=10)

        t_guide = self.tabview.add(TAB_GUIDE)
        t_ai = self.tabview.add(TAB_AI)
        t_tips = self.tabview.add(TAB_TIPS)
        t_faq = self.tabview.add(TAB_FAQ)
        t_about = self.tabview.add(TAB_ABOUT)

        self._build_guide_tab(t_guide)
        self._build_ai_tab(t_ai)
        self._build_tips_tab(t_tips)
        self._build_faq_tab(t_faq)
        self._build_about_tab(t_about)

        tab_map = {
            "guide": TAB_GUIDE,
            "ai": TAB_AI,
            "tips": TAB_TIPS,
            "faq": TAB_FAQ,
            "about": TAB_ABOUT,
        }
        self.tabview.set(tab_map.get(initial_tab, TAB_GUIDE))

    # ==============================
    # タブ1: 🚀 使い方ガイド
    # ==============================
    def _build_guide_tab(self, parent_tab):
        scroll = ctk.CTkScrollableFrame(parent_tab, fg_color=BG_COLOR)
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        self._section_heading(scroll, "かんたん4ステップ！基本的な使い方の流れ", pady=(6, 4))

        steps = [
            (
                "STEP 1",
                PRIMARY,
                "AI / API サービスを選んで接続設定を行う",
                "左上の「AI / API 設定」から使用するサービス（Gemini 無料 / Gemini 課金 / Azure OpenAI）を選び、「⚙️ AI詳細設定」ボタンを押します。\n"
                "APIキー（Azure OpenAIの場合はエンドポイントURL・APIキー・デプロイ名）を入力し、「🔌 接続テスト」で通信が正常か確認してから「設定を適用して保存・閉じる」を押してください。",
            ),
            (
                "STEP 2",
                COLOR_INFO,
                "図面ファイルと解析したい項目を選択する",
                "「ファイルを選択」ボタンから 2D 機械図面（PDF または PNG/JPEG/BMP 画像）を選びます。選択するとすぐに右上にプレビューが表示されます。\n"
                "複数ページの PDF は「対象ページ」から解析するページを選べます。続いて「解析項目の選択」で必要な情報（体積・重量・搬送 / 部材表・切断寸法 / 使用購入品 / 使用機械加工品）にチェックを入れ、必要に応じて「基本比重」と「基本単価 (円/kg)」を調整してください。",
            ),
            (
                "STEP 3",
                COLOR_SUCCESS,
                "「解析を実行」ボタンを押して結果を確認する",
                "「解析を実行」を押すと、図面の文字座標を画像に照合（Visual Grounding）したうえで AI が解析を行います。\n"
                "完了すると右側の5つのタブ（「体積・重量」「部材表 (切断寸法)」「使用購入品」「使用機械加工品」「見積り・搬送サマリー」）に結果が表示されます。途中で止めたい場合は「⏹ 解析を中止」ボタンで安全に中断できます。",
            ),
            (
                "STEP 4",
                COLOR_PURPLE,
                "単価調整・拡大確認・新規 Excel で開いて印刷 / 任意保存を行う",
                "「見積り・搬送サマリー」タブの「💴 単価設定・再計算」ボタンから材質ごとの単価(円/kg)を変更すると、AI再通信なしで概算材料費を即再計算できます。\n"
                "「📊 Excelで開く」ボタンを押すと、A4横・横1ページ収縮・見出し繰り返し・合計数式入りの新規Excelブックが直接開くため、そのまま印刷でき、必要な場合だけ任意の名前で保存できます。",
            ),
        ]
        for badge, color, title, desc in steps:
            self._create_accent_card(scroll, color, badge, title, desc)

        self._section_heading(scroll, "▼ 5つの結果タブと抽出・集計される内容")
        self._create_styled_table(
            scroll,
            ["タブ・項目", "どんな業務に使えるか", "主な出力内容"],
            [
                [
                    "📐 体積・重量",
                    "部品ごとの数量を反映した正確な総重量・材料費の見積り",
                    "品番、部品名/要素名、材質(比重)、数量、読み取り寸法、体積、単重、小計重量、計算式",
                ],
                [
                    "🧱 部材表 (切断寸法)",
                    "工場への切断指示・鋼材手配用の部材表（木取りリスト）作成",
                    "品番、部位、部材種別(PL/FB/L鋼等)、材質、断面規格・板厚、切断寸法、端部切断、数量、合計長さ、小計重量、表面処理",
                ],
                [
                    "🛒 使用購入品",
                    "ボルト・軸受・空圧機器などの市販品・規格品の発注手配",
                    "品番、品名、メーカー、型式・規格、材質・仕様、数量、単位(個/本/組等)、手配区分、備考",
                ],
                [
                    "⚙️ 使用機械加工品",
                    "切削・旋盤・フライス・研削等の製作部品・外注見積り手配",
                    "品番、部品名、材質、素材・概略寸法、加工工程・内容、公差・粗さ、熱処理・表面処理、数量",
                ],
                [
                    "🚛 見積り・搬送サマリー",
                    "概算材料費の算出・搬送トラック選定・鋼材定尺本数の確認",
                    "推奨トラック車種（重量＋外形寸法＋最大長尺から自動判定）、荷役目安、材質別の重量＆概算材料費（単価変更・再計算対応）、5.5m/6.0m定尺発注目安本数",
                ],
            ],
            [2, 3, 5],
        )

        self._section_heading(scroll, "▼ 便利な操作一覧")
        self._create_styled_table(
            scroll,
            ["操作場所", "操作方法", "できること"],
            [
                ["図面プレビュー", "クリック または「🔍 拡大表示」", "別ウィンドウで拡大ビューアーを開く（ホイールで拡大縮小、ドラッグで移動）"],
                ["結果テーブル（全タブ）", "表の上で右クリック", "「選択行をコピー」「表全体をコピー (Excel貼り付け用)」を実行"],
                ["部材 / 購入品 / 加工品 / 集計", "列の見出しをクリック", "クリックした列（品番・メーカー・材質など）で昇順 / 降順に並べ替え"],
                ["見積り・搬送サマリータブ", "「💴 単価設定・再計算」をクリック", "材質別単価 (円/kg) を変更し、AI再通信なしで概算材料費を即時再計算"],
                ["左パネル下部", "「📊 Excelで開く」をクリック", "印刷設定済みの新規Excelブックを直接開く（そのまま印刷・必要時のみ任意保存）"],
                ["対象ページ欄", "プルダウンからページを選択", "複数ページのPDFから任意のページをプレビュー・解析対象に切り替え"],
            ],
            [2, 3, 5],
        )

    # ==============================
    # タブ2: 🤖 AI・API の準備
    # ==============================
    def _build_ai_tab(self, parent_tab):
        scroll = ctk.CTkScrollableFrame(parent_tab, fg_color=BG_COLOR)
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        self._section_heading(scroll, "AIエンジンの準備と初期設定手順", pady=(6, 4))

        def _gemini_buttons(box):
            ctk.CTkButton(
                box,
                text="🌐 Google AI Studio (APIキー取得) を開く",
                fg_color=PRIMARY,
                hover_color=PRIMARY_HOVER,
                text_color="white",
                command=lambda: webbrowser.open_new("https://aistudio.google.com/app/apikey"),
            ).pack(side="left", padx=(0, 8))
            if self._on_open_api_settings:
                ctk.CTkButton(
                    box,
                    text="⚙️ AI詳細設定を開く",
                    fg_color="#E9ECEF",
                    hover_color="#DEE2E6",
                    text_color=TEXT_COLOR,
                    border_width=1,
                    border_color=BORDER_COLOR,
                    command=self._open_api_settings_and_close,
                ).pack(side="left")

        self._create_accent_card(
            scroll,
            PRIMARY,
            "推奨 1: Gemini API",
            "Google Gemini API を使う場合（個人・検証利用にも最適 / 無料枠あり）",
            "Google の最新マルチモーダル AI を使用して機械図面を解析します。無料枠（Free）と課金枠（Paid）を個別に登録して切り替えられます。\n\n"
            "【設定手順】\n"
            "1. 下のボタンから「Google AI Studio」へアクセスし、Google アカウントでログインします。\n"
            "2. 「Create API key（APIキーを作成）」を押して発行されたキーをコピーします。\n"
            "3. 「⚙️ AI詳細設定」の「Gemini 無料枠」または「Gemini 課金枠」タブの APIキー欄に右クリック→「貼り付け」します。\n"
            "4. 「🌐 最新モデルを更新」を押すと最新モデル一覧の取得と同時に RPM・スレッド・パラメータも自動最適化されます。\n"
            "5. 「🔌 接続テスト」で正常通信を確認し、「設定を適用して保存・閉じる」を押してください。",
            extra_builder=_gemini_buttons,
        )

        def _azure_buttons(box):
            ctk.CTkButton(
                box,
                text="🌐 Azure OpenAI 公式ドキュメントを開く",
                fg_color=COLOR_PURPLE,
                hover_color="#59339D",
                text_color="white",
                command=lambda: webbrowser.open_new(
                    "https://learn.microsoft.com/ja-jp/azure/ai-services/openai/"
                ),
            ).pack(side="left", padx=(0, 8))
            if self._on_open_api_settings:
                ctk.CTkButton(
                    box,
                    text="⚙️ AI詳細設定を開く",
                    fg_color="#E9ECEF",
                    hover_color="#DEE2E6",
                    text_color=TEXT_COLOR,
                    border_width=1,
                    border_color=BORDER_COLOR,
                    command=self._open_api_settings_and_close,
                ).pack(side="left")

        self._create_accent_card(
            scroll,
            COLOR_PURPLE,
            "推奨 2: Azure OpenAI",
            "Azure OpenAI Service を使う場合（法人・社外秘図面・セキュア環境向け）",
            "Microsoft Azure 上の企業向け OpenAI モデルを利用します。入力データが外部の AI 学習に利用されないため、機密性の高い機械図面でも安心して解析できます。\n\n"
            "【設定手順】\n"
            "1. Azure Portal または Azure AI Foundry でリソースの以下3点を確認します。\n"
            "   ・エンドポイントURL（例: https://xxxx.openai.azure.com/）※長いターゲットURI全体を貼っても自動補正されます\n"
            "   ・APIキー（KEY 1 または KEY 2）\n"
            "   ・デプロイ名（画像認識対応モデルのデプロイ名。全ステップ共通でもステップ別でも設定可）\n"
            "2. 「⚙️ AI詳細設定」→「Azure OpenAI Service」タブに入力し、「🌐 最新モデルを更新」を押します（次世代 v1 API を自動判定し、パラメータも最適化されます）。\n"
            "3. 「🔌 接続テスト」で成功を確認し、「設定を適用して保存・閉じる」を押してください。",
            extra_builder=_azure_buttons,
        )

        self._section_heading(scroll, "▼ Gemini API と Azure OpenAI Service の比較")
        self._create_styled_table(
            scroll,
            ["比較項目", "Gemini API (無料枠)", "Gemini API (課金枠)", "Azure OpenAI Service"],
            [
                ["利用料金", "完全無料（回数制限あり）", "従量課金", "Azure 従量課金（法人契約可）"],
                [
                    "データ学習への利用",
                    "Googleのモデル改善に利用される可能性あり",
                    "AI学習には利用されない",
                    "AI学習には利用されない（企業向け高セキュリティ）",
                ],
                [
                    "処理速度・並列実行",
                    "直列実行（スレッド 1）推奨",
                    "高速並列実行に対応",
                    "高速並列実行に対応（次世代 v1 API 対応）",
                ],
                [
                    "おすすめの用途",
                    "お試し・公開図面の検証",
                    "大量図面・高精度解析",
                    "社外秘図面・法人業務での本格運用",
                ],
            ],
            [2, 2, 3, 3],
        )

    # ==============================
    # タブ3: 📐 精度を上げるコツ
    # ==============================
    def _build_tips_tab(self, parent_tab):
        scroll = ctk.CTkScrollableFrame(parent_tab, fg_color=BG_COLOR)
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        self._section_heading(scroll, "図面解析の精度とスピードを高めるポイント", pady=(6, 4))

        self._create_accent_card(
            scroll,
            COLOR_SUCCESS,
            "図面データ",
            "CADから直接出力した「ベクターPDF」が最も高精度です",
            "・CADソフトから直接PDF出力した図面には文字の位置データが含まれており、本アプリが自動で「赤い枠と青いID番号（Visual Grounding）」を図面に重ねてAIへ渡すため、寸法値や部品表の読み違いを大幅に減らせます。\n"
            "・紙をスキャンしたPDFや画像ファイル（PNG/JPEG）も解析可能ですが、なるべく文字がつぶれていない鮮明な画像（200〜300dpi相当以上）をご用意ください。",
        )

        self._create_accent_card(
            scroll,
            PRIMARY,
            "モデル使い分け",
            "Step1〜2 は高速モデル、Step3 は高精度推論モデルがおすすめです",
            "・Step1（一覧抽出）と Step2（部品ごとの並列要素抽出）は処理回数が多くなるため、応答が速い主力モデル（Gemini 4 Hydrogen や gpt-5-mini 等）を選ぶと高速かつ安定します。\n"
            "・Step3（見落とし最終検証・使用部材/購入品/機械加工品の分類抽出）は図面全体の理解力と公差・加工工程の読解が重要なため、高精度モデル（Gemini 4 Neon や gpt-5 / gpt-6 等）を割り当てると精度が向上します。\n"
            "・「⚙️ AI詳細設定」の「🌐 最新モデルを更新」を押すと、選択モデルに合わせて RPM・スレッド・最大トークンが自動最適化されます。",
        )

        self._create_accent_card(
            scroll,
            COLOR_WARNING,
            "カスタムプロンプト",
            "社内ルールや図面固有の癖は「独自の追加指示」に登録しておけます",
            "「⚙️ AI詳細設定」下部の「独自の追加指示 (カスタムプロンプト)」に具体的なルールを登録し、よく使うものは「⭐ お気に入り」に保存しておくと非常に便利です。\n\n"
            "【登録例】\n"
            "・材質の記載がない部品は SS400（比重 7.85）として計算してください\n"
            "・図面右上の部品表に記載された数量を必ず読み取ってください\n"
            "・市販の締結部品（ボルト・ナット・座金）は必ず「使用購入品」に分類してください\n"
            "・溶接構造のフレームは「使用機械加工品」と「使用部材」の両方に明記してください",
        )

        self._create_accent_card(
            scroll,
            COLOR_DANGER,
            "重要なお願い",
            "AIによる計算・抽出結果は必ず目視でご確認ください",
            "・複雑な三面図の隠れ線解釈や、かすれた寸法値の読み取りには AI の推測が含まれる場合があります。\n"
            "・体積タブの「備考 (計算式等)」列にAIが組み立てた計算式（例: [式: 150*100*10]）が表示されるため、図面拡大ビューアーと見比べながら根拠を検算していただけます。",
        )

    # ==============================
    # タブ4: ❓ よくある質問 (FAQ)
    # ==============================
    def _build_faq_tab(self, parent_tab):
        scroll = ctk.CTkScrollableFrame(parent_tab, fg_color=BG_COLOR)
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        self._section_heading(scroll, "よくある質問・トラブル解決 (FAQ)", pady=(6, 4))

        faqs = [
            (
                "Q1. 接続テストや解析実行時にエラー（401 / 403 / 404 / 429 等）が表示されます",
                "エラー発生時に表示される診断ダイアログの「考えられる原因」と「対処方法」をご確認ください。\n"
                "・401（認証エラー）: APIキーのコピー漏れや前後の余計な文字、別リソースのキーとエンドポイントの組み合わせミスが主な原因です。\n"
                "・404（モデル/デプロイ未検出）: 「🌐 最新モデルを更新」を押して最新モデルを取得するか、Azure 側で作成した正確な「デプロイ名」を入力してください。\n"
                "・429（利用枠超過）: 無料枠の連続送信制限または1日の上限に達しています。少し時間を置くか、RPM・スレッド数を下げて再実行してください。",
            ),
            (
                "Q2. 解析に時間がかかります。もっと高速化する方法はありますか？",
                "以下の工夫で大幅に短縮できます。\n"
                "・左パネルの「解析項目の選択」で、今必要な項目だけにチェックを絞る（例: 部材表と購入品リストだけ欲しいときは「体積・重量・搬送」のチェックを外すと部品ごとの三面図分解をスキップするため非常に高速です。部材表から計算された重量で材料費・トラック目安も自動算出されます）。\n"
                "・Gemini 課金枠や Azure OpenAI をご利用の場合、スレッド数を 2〜5 に設定すると部品ごとの詳細解析や最終検証と部材表・手配リスト解析が並列で同時実行されます。",
            ),
            (
                "Q3. 複数ページあるPDFの2ページ目以降を解析したいです",
                "複数ページのPDFを選択すると、左パネルの「対象ページ」プルダウンが有効になります。\n"
                "解析したいページ番号（例: 2 / 5）を選ぶと右上のプレビューが即座に切り替わり、そのページを対象に解析を実行できます。",
            ),
            (
                "Q4. 図面の細かい寸法や、赤枠・青色のID番号をよく確認したいです",
                "右上の図面プレビュー画像をクリックするか、プレビュー右上の「🔍 拡大表示」ボタンを押してください。\n"
                "専用の拡大ビューアーが開き、マウスホイールでの拡大・縮小やドラッグでの移動ができます。",
            ),
            (
                "Q5. 解析結果をExcelで印刷したり、必要に応じてファイル保存したりしたいです",
                "・そのままExcelで確認・印刷したい場合: 左パネルの「📊 Excelで開く」ボタンを押してください。見積り・搬送サマリーや部材表・購入品手配表などを5シートに整理した新規の未保存Excelブックが開きます。\n"
                "・全シートにA4横・横1ページ自動収縮・2ページ目以降のタイトル/見出し繰り返し・ページ番号・合計(=SUM)数式が設定されているため、開いてすぐ印刷できます。残したい場合だけExcel上で「名前を付けて保存」してください。\n"
                "・一部だけコピーしたい場合: 結果の表の上で右クリックし、「選択行をコピー」または「表全体をコピー (Excel貼り付け用)」を選ぶとクリップボードへコピーできます。",
            ),
            (
                "Q6. 概算材料費の単価（円/kg）を自社の仕入れ単価に変更できますか？",
                "はい、変更できます。「見積り・搬送サマリー」タブの「💴 単価設定・再計算」ボタンを押すと、SS400・SUS304・A5052 など材質別の単価(円/kg)と基本単価を編集できます。\n"
                "「適用して再計算」を押すと AI へ再通信することなく即座に見積り金額が再計算され、変更した単価は次回起動時にも引き継がれます。",
            ),
            (
                "Q7. 解析を途中で止めたくなった場合はどうすればよいですか？",
                "解析実行中は「解析を実行」ボタンが赤い「⏹ 解析を中止」ボタンに変わります。\n"
                "クリックするとレート制限待機中やリトライ待機中でも安全に処理を中断して待機状態へ戻ります。",
            ),
        ]
        for q_title, a_body in faqs:
            self._create_accent_card(scroll, PRIMARY, "FAQ", q_title, f"【回答】\n{a_body}")

    # ==============================
    # タブ5: 📄 Readme / 情報
    # ==============================
    def _build_about_tab(self, parent_tab):
        scroll = ctk.CTkScrollableFrame(parent_tab, fg_color=BG_COLOR)
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        self._section_heading(scroll, "アプリケーション情報", pady=(6, 4))

        def _folder_btn(box):
            if self._settings_dir:
                ctk.CTkButton(
                    box,
                    text="📂 設定保存フォルダを開く",
                    fg_color="#E9ECEF",
                    hover_color="#DEE2E6",
                    text_color=TEXT_COLOR,
                    border_width=1,
                    border_color=BORDER_COLOR,
                    command=self._open_settings_folder,
                ).pack(side="left")

        info_text = (
            f"アプリ名: {APP_NAME}（2D機械図面 解析ツール）\n"
            f"バージョン: {APP_VERSION}\n"
            f"対応ファイル形式: PDF（複数ページ対応） / PNG / JPEG / BMP\n"
            f"対応AIエンジン: Google Gemini API（無料枠 / 課金枠）、Microsoft Azure OpenAI Service（v1 / クラシック）\n"
            f"設定ファイル保存先: {self._settings_dir or 'ユーザーフォルダ内 DrawingAnalysisMiya'}"
        )
        self._create_accent_card(
            scroll, PRIMARY, APP_VERSION, f"{APP_NAME} について", info_text, extra_builder=_folder_btn
        )

        self._section_heading(scroll, "▼ アプリ概要 (README)")
        readme_frame = ctk.CTkFrame(
            scroll,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        readme_frame.pack(fill="both", expand=True, padx=8, pady=(4, 12))

        readme_box = st.ScrolledText(
            readme_frame,
            wrap=tk.WORD,
            font=("Meiryo UI", 9),
            bg=CARD_BG,
            fg=TEXT_COLOR,
            relief=tk.FLAT,
            height=18,
            padx=14,
            pady=12,
        )
        readme_box.pack(fill="both", expand=True, padx=4, pady=4)
        readme_box.insert(tk.END, _load_readme_text())
        readme_box.configure(state=tk.DISABLED)
