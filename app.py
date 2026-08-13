# -*- coding: utf-8 -*-
"""
Minivibe（ミニバイブ）
--------------------------------------------------------------
冗長なプレゼン資料を「Apple発表会風ミニマリズム」に自動変換し、
アプリ内でリアルタイムに編集・並べ替えしたのち、
PowerPoint(.pptx)として出力できるStreamlitアプリ。

起動方法:
    pip install -r requirements.txt
    streamlit run app.py

必要なAPIキー（いずれか1つ以上）:
    - Cohere APIキー（メインエンジン・推奨）: https://dashboard.cohere.com/
    - Google Gemini APIキー（補助）        : https://aistudio.google.com/
    - Groq APIキー（補助）                 : https://console.groq.com/
"""

import io
import json
import re
import uuid
import html

import streamlit as st

# ページ設定は最初のStreamlitコマンドである必要がある
st.set_page_config(page_title="Minivibe ✨", page_icon="✨", layout="wide")

import pandas as pd

# python-pptxはアプリの核となる必須ライブラリのため、
# 未インストールの場合はここで分かりやすいエラーを出して停止する。
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError:
    st.error(
        "`python-pptx` がインストールされていません。\n\n"
        "ターミナルで `pip install python-pptx` を実行してから、アプリを再起動してください。"
    )
    st.stop()


# ============================================================
# 定数・テーマ定義
# ============================================================

THEMES = {
    "Apple Black": {
        "bg": "#000000",
        "text": "#FFFFFF",
        "sub": "#98989D",
        "metric": "#FFFFFF",
        "border": "#1D1D1F",
    },
    "Apple White": {
        "bg": "#FFFFFF",
        "text": "#1D1D1F",
        "sub": "#86868B",
        "metric": "#1D1D1F",
        "border": "#E5E5E7",
    },
    "Cyber": {
        "bg": "#05050A",
        "text": "#EAFBFF",
        "sub": "#7DFFEA",
        "metric": "#00FFC6",
        "border": "#1B1B2F",
    },
    "Minimal Grey": {
        "bg": "#F5F5F7",
        "text": "#1D1D1F",
        "sub": "#6E6E73",
        "metric": "#1D1D1F",
        "border": "#D2D2D7",
    },
}

FONT_STACK_CSS = (
    "-apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Hiragino Sans', "
    "'Yu Gothic', 'Segoe UI', Roboto, sans-serif"
)
PPTX_FONT_NAME = "Yu Gothic"  # 日本語表示に強いフォント（未インストール環境ではOS側で代替されます）

DEFAULT_COHERE_MODEL = "command-r-plus"
DEFAULT_GEMINI_MODEL = "gemini-2.0-flash"
DEFAULT_GROQ_MODEL = "llama-3.3-70b-versatile"


# ============================================================
# ユーティリティ
# ============================================================

def hexstr(h: str) -> str:
    """'#FFFFFF' -> 'FFFFFF' に変換（RGBColor.from_string用）"""
    return h.lstrip("#").upper()


def metric_scale(metric_text: str) -> float:
    """巨大数値の文字数に応じたスケール係数(0.0〜1.0)。長い文字列でもはみ出しにくくする。"""
    length = len(str(metric_text or "").strip())
    if length <= 4:
        return 1.0
    elif length <= 8:
        return 0.75
    elif length <= 14:
        return 0.5
    else:
        return 0.32


def new_slide_dict(headline: str = "", metric: str = "", subtext: str = "") -> dict:
    return {
        "id": str(uuid.uuid4()),
        "headline": headline,
        "metric": metric,
        "subtext": subtext,
    }


# ============================================================
# AIプロンプト生成
# ============================================================

def build_prompt(source_text: str, n_slides: int) -> str:
    instruction = (
        "あなたは優秀なエンジニアです。冗長な説明を排除し、"
        "AppleのKeynoteプレゼンテーションのような最小限かつ洗練された表現に変換すること。"
    )
    return f"""{instruction}

以下の原稿を、プレゼンテーション用に約 {n_slides} 枚のスライドへ分解してください。
各スライドは「1枚につき1メッセージ」の原則を厳守し、次の3要素のみで構成してください。

1. headline : 短く強烈な一文のメインコピー（日本語で15〜20文字程度）
2. metric   : 印象に残る数値やデータ（例: "10x", "99.9%", "3倍"。該当がなければ短いキーワードでも可）
3. subtext  : 補足説明（1〜2行、40文字程度まで）

【出力形式（厳守）】
説明文・前置き・Markdownのコードブロック(```)は一切使わず、次のJSON配列のみを出力してください。

[
  {{"headline": "...", "metric": "...", "subtext": "..."}},
  {{"headline": "...", "metric": "...", "subtext": "..."}}
]

【変換対象の原稿】
{source_text}
"""


def extract_json_array(raw_text: str):
    """AIの応答からJSON配列部分を頑健に抽出する"""
    if not raw_text or not raw_text.strip():
        raise ValueError("AIからの応答が空でした。")

    text = raw_text.strip()
    text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
    text = re.sub(r"```$", "", text).strip()

    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1 or end < start:
        raise ValueError(f"応答からJSON配列が見つかりませんでした。応答冒頭: {text[:200]}")

    json_str = text[start:end + 1]
    try:
        data = json.loads(json_str)
    except json.JSONDecodeError as e:
        raise ValueError(f"JSONの解析に失敗しました: {e}\n応答冒頭: {json_str[:300]}")

    if not isinstance(data, list):
        raise ValueError("AIの応答がリスト（配列）形式ではありませんでした。")
    return data


# ============================================================
# AI呼び出し（Cohereをメイン、Gemini / Groqを補助として実装）
# ============================================================

def call_cohere(prompt: str, api_key: str, model: str = DEFAULT_COHERE_MODEL) -> str:
    try:
        import cohere
    except ImportError:
        raise RuntimeError("`cohere` ライブラリが未インストールです。`pip install cohere` を実行してください。")

    last_err = None

    # 1) 新しいCohere SDK (ClientV2) を優先
    try:
        co = cohere.ClientV2(api_key=api_key)
        resp = co.chat(model=model, messages=[{"role": "user", "content": prompt}])
        content = getattr(getattr(resp, "message", None), "content", None)
        if isinstance(content, list) and len(content) > 0 and hasattr(content[0], "text"):
            return content[0].text
        if content:
            return str(content)
        return str(resp)
    except Exception as e:
        last_err = e

    # 2) 旧SDK (Client) にフォールバック
    try:
        co = cohere.Client(api_key)
        resp = co.chat(model=model, message=prompt)
        if hasattr(resp, "text"):
            return resp.text
        return str(resp)
    except Exception as e2:
        raise RuntimeError(
            f"Cohere APIの呼び出しに失敗しました。(ClientV2: {last_err} / Client: {e2})"
        )


def call_gemini(prompt: str, api_key: str, model: str = DEFAULT_GEMINI_MODEL) -> str:
    try:
        import google.generativeai as genai
    except ImportError:
        raise RuntimeError(
            "`google-generativeai` ライブラリが未インストールです。`pip install google-generativeai` を実行してください。"
        )

    try:
        genai.configure(api_key=api_key)
        gm = genai.GenerativeModel(model)
        resp = gm.generate_content(prompt)
        text = getattr(resp, "text", None)
        if text:
            return text
        if getattr(resp, "candidates", None):
            parts = resp.candidates[0].content.parts
            return "".join(getattr(p, "text", "") for p in parts)
        raise RuntimeError("Geminiから有効な応答が得られませんでした。")
    except Exception as e:
        raise RuntimeError(f"Gemini APIの呼び出しに失敗しました: {e}")


def call_groq(prompt: str, api_key: str, model: str = DEFAULT_GROQ_MODEL) -> str:
    try:
        from groq import Groq
    except ImportError:
        raise RuntimeError("`groq` ライブラリが未インストールです。`pip install groq` を実行してください。")

    try:
        client = Groq(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "あなたは優秀なエンジニアです。冗長な説明を排除し、"
                        "AppleのKeynoteプレゼンテーションのような最小限かつ洗練された表現に変換すること。"
                    ),
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.7,
        )
        return resp.choices[0].message.content
    except Exception as e:
        raise RuntimeError(f"Groq APIの呼び出しに失敗しました: {e}")


def generate_slides_with_ai(source_text, engine, cohere_key, gemini_key, groq_key,
                             cohere_model, gemini_model, groq_model, n_slides):
    prompt = build_prompt(source_text, n_slides)

    if engine == "Cohere":
        if not cohere_key:
            raise ValueError("Cohere APIキーが入力されていません。サイドバーに入力してください。")
        raw = call_cohere(prompt, cohere_key, cohere_model)
    elif engine == "Google Gemini":
        if not gemini_key:
            raise ValueError("Gemini APIキーが入力されていません。サイドバーに入力してください。")
        raw = call_gemini(prompt, gemini_key, gemini_model)
    elif engine == "Groq":
        if not groq_key:
            raise ValueError("Groq APIキーが入力されていません。サイドバーに入力してください。")
        raw = call_groq(prompt, groq_key, groq_model)
    else:
        raise ValueError(f"不明なエンジンです: {engine}")

    st.session_state["_last_ai_raw"] = raw

    data = extract_json_array(raw)
    slides = []
    for item in data:
        if not isinstance(item, dict):
            continue
        slides.append(new_slide_dict(
            headline=str(item.get("headline", "")).strip(),
            metric=str(item.get("metric", "")).strip(),
            subtext=str(item.get("subtext", "")).strip(),
        ))

    if not slides:
        raise ValueError("AIの応答からスライドを1件も抽出できませんでした。")
    return slides


# ============================================================
# スライド操作（st.session_state管理）
# ============================================================

def move_slide(index: int, direction: int):
    slides = st.session_state.slides
    new_index = index + direction
    if 0 <= new_index < len(slides):
        slides[index], slides[new_index] = slides[new_index], slides[index]


def delete_slide(index: int):
    slides = st.session_state.slides
    if 0 <= index < len(slides):
        slides.pop(index)


def add_blank_slide():
    st.session_state.slides.append(
        new_slide_dict(headline="新しいメッセージ", metric="0", subtext="補足テキストを入力")
    )


# ============================================================
# プレビューカード（HTML/CSS）
# ============================================================

def render_preview_card_html(slide: dict, theme: dict) -> str:
    headline = html.escape(slide.get("headline", "") or "")
    metric = html.escape(slide.get("metric", "") or "")
    subtext = html.escape(slide.get("subtext", "") or "")

    scale = metric_scale(slide.get("metric", ""))
    metric_px = int(96 * scale)

    return f"""
    <div style="
        background:{theme['bg']};
        color:{theme['text']};
        border:1px solid {theme['border']};
        border-radius:24px;
        padding:44px 32px;
        margin-bottom:18px;
        text-align:center;
        box-shadow:0 10px 30px rgba(0,0,0,0.18);
        font-family:{FONT_STACK_CSS};
    ">
        <div style="
            font-size:19px; font-weight:700; letter-spacing:0.3px;
            color:{theme['text']}; margin-bottom:16px; opacity:0.92;
        ">{headline}</div>
        <div style="
            font-size:{metric_px}px; font-weight:800; line-height:1.05;
            color:{theme['metric']}; margin:6px 0 16px 0; word-break:break-word;
        ">{metric}</div>
        <div style="
            font-size:14px; font-weight:400; color:{theme['sub']};
        ">{subtext}</div>
    </div>
    """


# ============================================================
# PowerPoint (.pptx) 生成
# ============================================================

def build_pptx(slides: list, theme_name: str) -> io.BytesIO:
    theme = THEMES[theme_name]

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    bg_color = RGBColor.from_string(hexstr(theme["bg"]))
    text_color = RGBColor.from_string(hexstr(theme["text"]))
    sub_color = RGBColor.from_string(hexstr(theme["sub"]))
    metric_color = RGBColor.from_string(hexstr(theme["metric"]))

    for slide_data in slides:
        slide = prs.slides.add_slide(blank_layout)

        # 背景色
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = bg_color

        # headline（上部・キーフレーズ）
        headline_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.3), Inches(11.33), Inches(1.0))
        tf = headline_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = slide_data.get("headline", "")
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.name = PPTX_FONT_NAME
        run.font.color.rgb = text_color

        # metric（中央・巨大数値）
        scale = metric_scale(slide_data.get("metric", ""))
        metric_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.3), Inches(12.33), Inches(2.9))
        tf2 = metric_box.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.CENTER
        run2 = p2.add_run()
        run2.text = slide_data.get("metric", "")
        run2.font.size = Pt(max(int(160 * scale), 40))
        run2.font.bold = True
        run2.font.name = PPTX_FONT_NAME
        run2.font.color.rgb = metric_color

        # subtext（下部・補足）
        sub_box = slide.shapes.add_textbox(Inches(2.0), Inches(5.5), Inches(9.33), Inches(1.3))
        tf3 = sub_box.text_frame
        tf3.word_wrap = True
        p3 = tf3.paragraphs[0]
        p3.alignment = PP_ALIGN.CENTER
        run3 = p3.add_run()
        run3.text = slide_data.get("subtext", "")
        run3.font.size = Pt(18)
        run3.font.name = PPTX_FONT_NAME
        run3.font.color.rgb = sub_color

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


# ============================================================
# セッション状態の初期化
# ============================================================

if "slides" not in st.session_state:
    st.session_state.slides = [
        new_slide_dict(
            headline="shorter. faster. bolder.",
            metric="10x",
            subtext="圧倒的な体験を、より少ない言葉で。",
        )
    ]


# ============================================================
# サイドバー UI
# ============================================================

with st.sidebar:
    st.header("⚙️ 設定")

    st.subheader("🔑 APIキー")
    st.caption("APIキーはこのセッション内でのみ使用され、サーバーに保存されません。")

    cohere_api_key = st.text_input("Cohere APIキー（メイン・推奨）", type="password", key="cohere_api_key")

    with st.expander("補助エンジン（Google Gemini / Groq）"):
        gemini_api_key = st.text_input("Google Gemini APIキー", type="password", key="gemini_api_key")
        groq_api_key = st.text_input("Groq APIキー", type="password", key="groq_api_key")

    engine = st.selectbox("🤖 生成に使うAIエンジン", ["Cohere", "Google Gemini", "Groq"], index=0)

    st.subheader("🎨 デザインテーマ")
    theme_name = st.selectbox("テーマを選択", list(THEMES.keys()), index=0, label_visibility="collapsed")

    with st.expander("詳細設定（モデル名 / デバッグ）"):
        cohere_model = st.text_input("Cohereモデル", value=DEFAULT_COHERE_MODEL)
        gemini_model = st.text_input("Geminiモデル", value=DEFAULT_GEMINI_MODEL)
        groq_model = st.text_input("Groqモデル", value=DEFAULT_GROQ_MODEL)
        show_debug = st.checkbox("AIの生応答を表示する", value=False)


# ============================================================
# メイン UI
# ============================================================

st.title("✨ Minivibe")
st.caption("冗長なプレゼン資料を、Apple発表会風ミニマリズムへ。")

with st.expander("📄 原稿からAIでスライドを自動生成", expanded=True):
    raw_text = st.text_area(
        "変換したい原稿・長文を貼り付けてください",
        height=180,
        placeholder="ここに元原稿を入力...（例：製品の特長を説明した長文など）",
    )
    gen_col1, gen_col2 = st.columns([3, 1])
    with gen_col2:
        n_slides = st.number_input("目安枚数", min_value=1, max_value=15, value=5, step=1)
    with gen_col1:
        st.write("")
        generate_clicked = st.button("🚀 AIでスライド生成", type="primary", use_container_width=True)

    if generate_clicked:
        if not raw_text or not raw_text.strip():
            st.warning("原稿を入力してください。")
        else:
            with st.spinner("AIが原稿を削ぎ落としています..."):
                try:
                    new_slides = generate_slides_with_ai(
                        raw_text, engine,
                        cohere_api_key, gemini_api_key, groq_api_key,
                        cohere_model, gemini_model, groq_model,
                        n_slides,
                    )
                    st.session_state.slides = new_slides
                    st.success(f"{len(new_slides)}枚のスライドを生成しました。")
                    st.rerun()
                except Exception as e:
                    st.error(f"生成に失敗しました: {e}")

    if show_debug and st.session_state.get("_last_ai_raw"):
        with st.expander("🔍 AIの生応答（デバッグ）"):
            st.code(st.session_state["_last_ai_raw"])

st.divider()

left_col, right_col = st.columns([1, 1])

# ---------------- 左カラム：編集エリア（アコーディオン） ----------------
with left_col:
    st.subheader("📝 編集")

    if st.button("➕ 新しいスライドを追加", use_container_width=True):
        add_blank_slide()
        st.rerun()

    if not st.session_state.slides:
        st.info("まだスライドがありません。AIで生成するか、上のボタンで追加してください。")
    else:
        for i, slide in enumerate(st.session_state.slides):
            sid = slide["id"]
            label = slide["headline"].strip() or "(未入力)"
            with st.expander(f"スライド {i + 1}：{label[:20]}", expanded=(i == 0)):
                slide["headline"] = st.text_input(
                    "メインコピー（短く強烈な一文）", value=slide["headline"], key=f"headline_{sid}"
                )
                slide["metric"] = st.text_input(
                    "強調数値（例: 10x, 99.9%）", value=slide["metric"], key=f"metric_{sid}"
                )
                slide["subtext"] = st.text_area(
                    "補足テキスト（1〜2行）", value=slide["subtext"], key=f"subtext_{sid}", height=70
                )

                b1, b2, b3 = st.columns(3)
                if b1.button("⬆️ 上へ", key=f"up_{sid}", use_container_width=True):
                    move_slide(i, -1)
                    st.rerun()
                if b2.button("⬇️ 下へ", key=f"down_{sid}", use_container_width=True):
                    move_slide(i, 1)
                    st.rerun()
                if b3.button("🗑️ 削除", key=f"del_{sid}", use_container_width=True):
                    delete_slide(i)
                    st.rerun()

        with st.expander("📋 全体構成一覧（テーブル表示）"):
            df = pd.DataFrame(st.session_state.slides)[["headline", "metric", "subtext"]]
            df.index = df.index + 1
            df.columns = ["メインコピー", "数値", "補足"]
            st.dataframe(df, use_container_width=True)

# ---------------- 右カラム：プレビューエリア ----------------
with right_col:
    st.subheader("👀 プレビュー")

    theme = THEMES[theme_name]

    if not st.session_state.slides:
        st.info("プレビューするスライドがありません。")
    else:
        for slide in st.session_state.slides:
            st.markdown(render_preview_card_html(slide, theme), unsafe_allow_html=True)

        try:
            pptx_buffer = build_pptx(st.session_state.slides, theme_name)
            st.download_button(
                "⬇️ PowerPointをダウンロード（.pptx）",
                data=pptx_buffer,
                file_name="minivibe_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary",
            )
        except Exception as e:
            st.error(f"PowerPointの生成に失敗しました: {e}")
