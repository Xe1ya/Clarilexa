# -*- coding: utf-8 -*-
"""
Clarilex（旧Minivibe）
--------------------------------------------------------------
冗長なプレゼン資料を「Apple発表会風ミニマリズム」に自動変換し、
アプリ内でリアルタイムに編集・並べ替えしたのち、
PowerPoint(.pptx)として出力できるStreamlitアプリ。

AIがスライドの文脈から画像検索キーワードを自動生成し、
Unsplash（APIキー設定時）またはプレースホルダー画像を
各スライドへ自動的に挿入する機能つき。

起動方法:
    pip install -r requirements.txt
    streamlit run app.py

必要なAPIキー:
    - Cohere / Google Gemini / Groq のいずれか1つ以上（スライド生成用）
        Cohere : https://dashboard.cohere.com/
        Gemini : https://aistudio.google.com/
        Groq   : https://console.groq.com/
    - Unsplash Access Key（任意・画像自動挿入の精度向上用）
        https://unsplash.com/developers
        未設定でも動作しますが、その場合はテーマに関係しないダミー
        プレースホルダー画像（picsum.photos）が使われます。
"""

import io
import json
import re
import uuid
import html
from urllib.parse import quote

import streamlit as st

# ページ設定は最初のStreamlitコマンドである必要がある
st.set_page_config(page_title="Clarilex ✨", page_icon="✨", layout="wide")

import pandas as pd

# python-pptx / Pillow はアプリの核となる必須ライブラリのため、
# 未インストールの場合はここで分かりやすいエラーを出して停止する。
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from PIL import Image as PILImage
except ImportError:
    st.error(
        "`python-pptx` または `Pillow` がインストールされていません。\n\n"
        "ターミナルで `pip install python-pptx Pillow` を実行してから、アプリを再起動してください。"
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

DEFAULT_COHERE_MODEL = "command-a-03-2025"   # 2026年8月時点のCohere主力チャットモデル
DEFAULT_GEMINI_MODEL = "gemini-3.6-flash"    # 2026年8月時点のGemini現行Flashモデル
DEFAULT_GROQ_MODEL = "openai/gpt-oss-120b"   # llama-3.3-70b-versatile廃止に伴う後継モデル

UNSPLASH_SEARCH_URL = "https://api.unsplash.com/search/photos"


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


def new_slide_dict(headline: str = "", metric: str = "", subtext: str = "",
                    image_keyword: str = "", image_url: str = "",
                    image_credit: str = "", show_image: bool = True) -> dict:
    return {
        "id": str(uuid.uuid4()),
        "headline": headline,
        "metric": metric,
        "subtext": subtext,
        "image_keyword": image_keyword,
        "image_url": image_url,
        "image_credit": image_credit,
        "show_image": show_image,
    }


# ============================================================
# AIプロンプト生成
# ============================================================

def build_prompt(source_text: str, n_slides: int, image_lang: str = "English (推奨)") -> str:
    instruction = (
        "あなたは優秀なエンジニアです。冗長な説明を排除し、"
        "AppleのKeynoteプレゼンテーションのような最小限かつ洗練された表現に変換すること。"
    )

    if image_lang.startswith("English"):
        image_lang_note = (
            "英語で2〜5単語程度。Unsplash等の海外ストックフォトサービスでの検索精度を優先すること。"
        )
    else:
        image_lang_note = "日本語で2〜5単語程度の具体的なキーワード。"

    return f"""{instruction}

以下の原稿を、プレゼンテーション用に約 {n_slides} 枚のスライドへ分解してください。
各スライドは「1枚につき1メッセージ」の原則を厳守し、次の4要素のみで構成してください。

1. headline      : 短く強烈な一文のメインコピー（日本語で15〜20文字程度）
2. metric        : 印象に残る数値やデータ（例: "10x", "99.9%", "3倍"。該当がなければ短いキーワードでも可）
3. subtext       : 補足説明（1〜2行、40文字程度まで）
4. image_keyword : スライドの文脈・テーマを象徴する、具体的で視覚的な画像検索キーワード。
                   {image_lang_note}
                   抽象的な概念語ではなく、実際に写真として存在しうる具体物・情景を指定すること。
                   例）テーマが「ハンムラビ法典」なら "ancient stone law tablet"、
                       テーマが「裁判」なら "modern courtroom trial" のように。

【出力形式（厳守）】
説明文・前置き・Markdownのコードブロック(```)は一切使わず、次のJSON配列のみを出力してください。

[
  {{"headline": "...", "metric": "...", "subtext": "...", "image_keyword": "..."}},
  {{"headline": "...", "metric": "...", "subtext": "...", "image_keyword": "..."}}
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
                             cohere_model, gemini_model, groq_model, n_slides,
                             image_lang="English (推奨)", auto_fetch_images=True,
                             unsplash_key=None):
    """AIでスライド本文を生成し、必要なら画像URLも自動取得して返す。
    戻り値: (slides: list[dict], image_warnings: list[str])
    """
    prompt = build_prompt(source_text, n_slides, image_lang)

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
            image_keyword=str(item.get("image_keyword", "")).strip(),
        ))

    if not slides:
        raise ValueError("AIの応答からスライドを1件も抽出できませんでした。")

    image_warnings = []
    if auto_fetch_images:
        for s in slides:
            kw = s.get("image_keyword", "").strip()
            if not kw:
                continue
            try:
                url, credit = fetch_image_url(kw, unsplash_key)
                s["image_url"] = url
                s["image_credit"] = credit
            except Exception as e:
                image_warnings.append(f"「{(s['headline'] or kw)[:15]}」: {e}")

    return slides, image_warnings


# ============================================================
# 画像検索・取得（Unsplash APIをメイン、未設定時はプレースホルダー）
# ============================================================

def _fetch_from_unsplash(keyword: str, access_key: str, timeout: int = 10):
    try:
        import requests
    except ImportError:
        raise RuntimeError("`requests` ライブラリが未インストールです。`pip install requests` を実行してください。")

    resp = requests.get(
        UNSPLASH_SEARCH_URL,
        params={"query": keyword, "per_page": 1, "orientation": "landscape"},
        headers={"Authorization": f"Client-ID {access_key}"},
        timeout=timeout,
    )
    resp.raise_for_status()
    data = resp.json()
    results = data.get("results") or []
    if not results:
        raise RuntimeError(f"Unsplashで '{keyword}' に一致する画像が見つかりませんでした。")

    photo = results[0]
    url = (photo.get("urls") or {}).get("regular") or (photo.get("urls") or {}).get("small")
    if not url:
        raise RuntimeError("Unsplashの応答に画像URLが含まれていませんでした。")

    photographer = (photo.get("user") or {}).get("name", "Unsplash")
    credit = f"Photo by {photographer} on Unsplash"
    return url, credit


def fetch_image_url(keyword: str, unsplash_access_key: str = None):
    """画像URLと（あれば）クレジット表記を返す。
    Unsplashキーがある場合はUnsplash検索を優先し、失敗時や未設定時は
    ダミープレースホルダー画像（picsum.photos、キーワードに応じて決定的に選ばれる）にフォールバックする。
    """
    keyword = (keyword or "").strip()
    if not keyword:
        raise ValueError("画像検索キーワードが空です。")

    if unsplash_access_key:
        try:
            return _fetch_from_unsplash(keyword, unsplash_access_key)
        except Exception:
            pass  # フォールバックへ

    seed = quote(keyword)
    return f"https://picsum.photos/seed/{seed}/900/600", None


@st.cache_data(show_spinner=False, ttl=3600)
def download_image_bytes(url: str) -> bytes:
    """pptx埋め込み用に画像バイナリを取得する。同一URLはセッション内でキャッシュされる。"""
    try:
        import requests
    except ImportError:
        raise RuntimeError("`requests` ライブラリが未インストールです。`pip install requests` を実行してください。")

    resp = requests.get(url, timeout=15)
    resp.raise_for_status()
    return resp.content


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

    has_image = bool(slide.get("show_image")) and bool(slide.get("image_url"))
    scale = metric_scale(slide.get("metric", ""))

    if has_image:
        image_url = html.escape(slide.get("image_url") or "")
        credit = slide.get("image_credit")
        credit_html = (
            f'<div style="font-size:10px; color:{theme["sub"]}; opacity:0.75; '
            f'margin-top:6px; text-align:right;">{html.escape(credit)}</div>'
            if credit else ""
        )
        metric_px = int(72 * scale)
        return f"""
        <div style="
            background:{theme['bg']};
            color:{theme['text']};
            border:1px solid {theme['border']};
            border-radius:24px;
            padding:36px;
            margin-bottom:18px;
            box-shadow:0 10px 30px rgba(0,0,0,0.18);
            font-family:{FONT_STACK_CSS};
            display:flex;
            align-items:center;
            gap:28px;
        ">
            <div style="flex:1; min-width:0; text-align:left;">
                <div style="
                    font-size:17px; font-weight:700; letter-spacing:0.3px;
                    color:{theme['text']}; margin-bottom:14px; opacity:0.92;
                ">{headline}</div>
                <div style="
                    font-size:{metric_px}px; font-weight:800; line-height:1.05;
                    color:{theme['metric']}; margin:4px 0 12px 0; word-break:break-word;
                ">{metric}</div>
                <div style="
                    font-size:13px; font-weight:400; color:{theme['sub']};
                ">{subtext}</div>
            </div>
            <div style="flex:1; min-width:0;">
                <img src="{image_url}" style="
                    width:100%; height:180px; object-fit:cover;
                    border-radius:16px; display:block;
                " />
                {credit_html}
            </div>
        </div>
        """

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

def add_cover_image(slide, image_bytes: bytes, left_in: float, top_in: float,
                     width_in: float, height_in: float):
    """CSSのobject-fit:coverのように、アスペクト比を保ったまま指定枠いっぱいに
    画像をクロップ配置する（python-pptxのcrop_*プロパティを利用）。"""
    img_w = img_h = None
    try:
        with PILImage.open(io.BytesIO(image_bytes)) as im:
            img_w, img_h = im.size
    except Exception:
        pass

    pic = slide.shapes.add_picture(
        io.BytesIO(image_bytes), Inches(left_in), Inches(top_in),
        width=Inches(width_in), height=Inches(height_in),
    )

    if img_w and img_h:
        target_ratio = width_in / height_in
        img_ratio = img_w / img_h
        if img_ratio > target_ratio:
            visible_w = target_ratio / img_ratio
            crop = (1 - visible_w) / 2
            pic.crop_left = crop
            pic.crop_right = crop
        elif img_ratio < target_ratio:
            visible_h = img_ratio / target_ratio
            crop = (1 - visible_h) / 2
            pic.crop_top = crop
            pic.crop_bottom = crop

    return pic


def build_pptx(slides: list, theme_name: str):
    """編集後のスライドデータから16:9のPowerPointを生成する。
    戻り値: (buffer: io.BytesIO, warnings: list[str])
    """
    theme = THEMES[theme_name]

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    bg_color = RGBColor.from_string(hexstr(theme["bg"]))
    text_color = RGBColor.from_string(hexstr(theme["text"]))
    sub_color = RGBColor.from_string(hexstr(theme["sub"]))
    metric_color = RGBColor.from_string(hexstr(theme["metric"]))

    warnings = []

    for slide_data in slides:
        slide = prs.slides.add_slide(blank_layout)

        # 背景色
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = bg_color

        headline_text = slide_data.get("headline", "")
        short_label = (headline_text or slide_data.get("image_keyword") or "")[:15]

        has_image = bool(slide_data.get("show_image")) and bool(slide_data.get("image_url"))
        image_bytes = None
        if has_image:
            try:
                image_bytes = download_image_bytes(slide_data["image_url"])
            except Exception as e:
                warnings.append(f"「{short_label}」の画像取得に失敗しました: {e}")
                has_image = False

        if has_image:
            # ---- 左：テキスト / 右：画像 の2カラムレイアウト ----
            text_left, text_width = 0.6, 6.0
            img_left, img_top, img_width, img_height = 6.9, 0.6, 5.83, 6.3

            headline_box = slide.shapes.add_textbox(Inches(text_left), Inches(1.4), Inches(text_width), Inches(1.1))
            tf = headline_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = headline_text
            run.font.size = Pt(28)
            run.font.bold = True
            run.font.name = PPTX_FONT_NAME
            run.font.color.rgb = text_color

            scale = metric_scale(slide_data.get("metric", ""))
            metric_box = slide.shapes.add_textbox(Inches(text_left), Inches(2.6), Inches(text_width), Inches(2.2))
            tf2 = metric_box.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.alignment = PP_ALIGN.LEFT
            run2 = p2.add_run()
            run2.text = slide_data.get("metric", "")
            run2.font.size = Pt(max(int(110 * scale), 32))
            run2.font.bold = True
            run2.font.name = PPTX_FONT_NAME
            run2.font.color.rgb = metric_color

            sub_box = slide.shapes.add_textbox(Inches(text_left), Inches(5.0), Inches(text_width), Inches(1.3))
            tf3 = sub_box.text_frame
            tf3.word_wrap = True
            p3 = tf3.paragraphs[0]
            p3.alignment = PP_ALIGN.LEFT
            run3 = p3.add_run()
            run3.text = slide_data.get("subtext", "")
            run3.font.size = Pt(16)
            run3.font.name = PPTX_FONT_NAME
            run3.font.color.rgb = sub_color

            try:
                add_cover_image(slide, image_bytes, img_left, img_top, img_width, img_height)
            except Exception as e:
                warnings.append(f"「{short_label}」の画像配置に失敗しました: {e}")

            credit = slide_data.get("image_credit")
            if credit:
                credit_box = slide.shapes.add_textbox(
                    Inches(img_left), Inches(img_top + img_height + 0.05), Inches(img_width), Inches(0.3)
                )
                tfc = credit_box.text_frame
                pc = tfc.paragraphs[0]
                pc.alignment = PP_ALIGN.RIGHT
                runc = pc.add_run()
                runc.text = credit
                runc.font.size = Pt(9)
                runc.font.name = PPTX_FONT_NAME
                runc.font.color.rgb = sub_color

        else:
            # ---- 画像なし：中央揃えの元レイアウト ----
            headline_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.3), Inches(11.33), Inches(1.0))
            tf = headline_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = headline_text
            run.font.size = Pt(32)
            run.font.bold = True
            run.font.name = PPTX_FONT_NAME
            run.font.color.rgb = text_color

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
    return buffer, warnings


# ============================================================
# セッション状態の初期化
# ============================================================

if "slides" not in st.session_state:
    st.session_state.slides = [
        new_slide_dict(
            headline="shorter. faster. bolder.",
            metric="10x",
            subtext="圧倒的な体験を、より少ない言葉で。",
            image_keyword="minimalist product design",
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

    st.subheader("🖼️ 画像検索（任意）")
    unsplash_api_key = st.text_input(
        "Unsplash Access Key", type="password", key="unsplash_api_key",
        help="https://unsplash.com/developers で無料取得できます。",
    )
    st.caption("未入力の場合は、キーワードに応じたダミープレースホルダー画像（picsum.photos）が使われます。")

    engine = st.selectbox("🤖 生成に使うAIエンジン", ["Cohere", "Google Gemini", "Groq"], index=0)

    st.subheader("🎨 デザインテーマ")
    theme_name = st.selectbox("テーマを選択", list(THEMES.keys()), index=0, label_visibility="collapsed")

    with st.expander("詳細設定（モデル名 / 画像 / デバッグ）"):
        cohere_model = st.text_input("Cohereモデル", value=DEFAULT_COHERE_MODEL)
        gemini_model = st.text_input("Geminiモデル", value=DEFAULT_GEMINI_MODEL)
        groq_model = st.text_input("Groqモデル", value=DEFAULT_GROQ_MODEL)
        image_lang = st.selectbox("画像検索キーワードの言語", ["English (推奨)", "日本語"], index=0)
        show_debug = st.checkbox("AIの生応答を表示する", value=False)


# ============================================================
# メイン UI
# ============================================================

st.title("✨ Clarilex")
st.caption("冗長なプレゼン資料を、Apple発表会風ミニマリズムへ。（旧Minivibe）")

with st.expander("📄 原稿からAIでスライドを自動生成", expanded=True):
    raw_text = st.text_area(
        "変換したい原稿・長文を貼り付けてください",
        height=180,
        placeholder="ここに元原稿を入力...（例：ハンムラビ法典から近代法の成立までを説明した長文など）",
    )
    gen_col1, gen_col2 = st.columns([3, 1])
    with gen_col2:
        n_slides = st.number_input("目安枚数", min_value=1, max_value=15, value=5, step=1)
    with gen_col1:
        st.write("")
        generate_clicked = st.button("🚀 AIでスライド生成", type="primary", use_container_width=True)

    auto_fetch_images = st.checkbox("🖼️ 生成と同時に画像も自動取得する", value=True)

    if generate_clicked:
        if not raw_text or not raw_text.strip():
            st.warning("原稿を入力してください。")
        else:
            with st.spinner("AIが原稿を削ぎ落とし、画像を選定しています..."):
                try:
                    new_slides, img_warnings = generate_slides_with_ai(
                        raw_text, engine,
                        cohere_api_key, gemini_api_key, groq_api_key,
                        cohere_model, gemini_model, groq_model,
                        n_slides, image_lang, auto_fetch_images, unsplash_api_key,
                    )
                    st.session_state.slides = new_slides
                    st.success(f"{len(new_slides)}枚のスライドを生成しました。")
                    if img_warnings:
                        st.warning("一部のスライドで画像取得に失敗しました：\n" + "\n".join(img_warnings))
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

                st.markdown("**🖼️ 画像**")
                img_col1, img_col2 = st.columns([2, 1])
                with img_col1:
                    slide["image_keyword"] = st.text_input(
                        "画像検索キーワード", value=slide.get("image_keyword", ""),
                        key=f"imgkw_{sid}", label_visibility="collapsed",
                        placeholder="例: ancient stone law tablet",
                    )
                with img_col2:
                    slide["show_image"] = st.checkbox(
                        "表示する", value=slide.get("show_image", True), key=f"showimg_{sid}"
                    )

                if st.button("🔄 画像を検索/更新", key=f"fetchimg_{sid}", use_container_width=True):
                    kw = (slide["image_keyword"] or "").strip()
                    if not kw:
                        st.warning("画像検索キーワードを入力してください。")
                    else:
                        with st.spinner("画像を検索中..."):
                            try:
                                url, credit = fetch_image_url(kw, unsplash_api_key)
                                slide["image_url"] = url
                                slide["image_credit"] = credit
                                st.rerun()
                            except Exception as e:
                                st.error(f"画像の取得に失敗しました: {e}")

                if slide.get("image_url"):
                    st.image(slide["image_url"], caption=slide.get("image_credit") or None, use_container_width=True)

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
            df = pd.DataFrame(st.session_state.slides)[["headline", "metric", "subtext", "image_keyword"]]
            df.index = df.index + 1
            df.columns = ["メインコピー", "数値", "補足", "画像キーワード"]
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
            pptx_buffer, pptx_warnings = build_pptx(st.session_state.slides, theme_name)
            if pptx_warnings:
                st.warning("PowerPoint生成時に一部の画像を取得できませんでした：\n" + "\n".join(pptx_warnings))
            st.download_button(
                "⬇️ PowerPointをダウンロード（.pptx）",
                data=pptx_buffer,
                file_name="clarilex_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary",
            )
        except Exception as e:
            st.error(f"PowerPointの生成に失敗しました: {e}")
