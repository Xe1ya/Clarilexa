import json
import uuid
from io import BytesIO
import streamlit as st
import pandas as pd
import cohere
import google.generativeai as genai
from groq import Groq
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# --- ページ基本設定 ---
st.set_page_config(page_title="Minivibe - Apple風スライド生成", layout="wide")

# --- セッション状態の初期化＆古いデータの自動補正 ---
if "slides" not in st.session_state:
    st.session_state.slides = []

# 古いセッション状態に残っている非辞書型データを強制クリーンアップ
cleaned_slides = []
for item in st.session_state.slides:
    if isinstance(item, dict):
        if "id" not in item:
            item["id"] = str(uuid.uuid4())
        cleaned_slides.append(item)
    elif isinstance(item, str):
        cleaned_slides.append({
            "id": str(uuid.uuid4()),
            "title": item,
            "number": "",
            "subtitle": ""
        })
st.session_state.slides = cleaned_slides


# --- AI生成関数 ---
def generate_slides_with_ai(text, cohere_key, gemini_key, groq_key):
    prompt = f"""あなたは優秀なエンジニア・プレゼンデザイナーです。
以下の入力文章を、Apple発表会風のミニマリズムなスライド（1スライド1メッセージ）に分解・変換してください。

【入力文章】
{text}

【出力フォーマット】
必ず以下の構造を持つJSON配列（リスト）形式のみで出力してください。Markdownの解説や余計なテキストは一切含めないでください。

[
  {{
    "title": "メインコピー（短く強烈な1文）",
    "number": "補足データ・強調数値（例: 10x, 99.9%, 2026 など。無ければ空文字）",
    "subtitle": "補足テキスト（1〜2行程度）"
  }}
]
"""

    raw_response = ""
    
    if cohere_key:
        try:
            co = cohere.Client(cohere_key)
            res = co.chat(
                model="command-r-plus-08-2024",
                message=prompt
            )
            raw_response = res.text
        except Exception as e:
            st.error(f"Cohere API エラー: {e}")
            return None
    elif gemini_key:
        try:
            genai.configure(api_key=gemini_key)
            model = genai.GenerativeModel('gemini-2.0-flash')
            res = model.generate_content(prompt)
            raw_response = res.text
        except Exception as e:
            st.error(f"Gemini API エラー: {e}")
            return None
    elif groq_key:
        try:
            client = Groq(api_key=groq_key)
            res = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": prompt}]
            )
            raw_response = res.choices[0].message.content
        except Exception as e:
            st.error(f"Groq API エラー: {e}")
            return None
    else:
        st.error("APIキーを入力してください。")
        return None

    try:
        clean_text = raw_response.strip()
        if clean_text.startswith("```json"):
            clean_text = clean_text[7:]
        if clean_text.startswith("```"):
            clean_text = clean_text[3:]
        if clean_text.endswith("```"):
            clean_text = clean_text[:-3]
            
        data = json.loads(clean_text.strip())
        
        normalized_slides = []
        if isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    normalized_slides.append({
                        "id": str(uuid.uuid4()),
                        "title": str(item.get("title", "")),
                        "number": str(item.get("number", "")),
                        "subtitle": str(item.get("subtitle", ""))
                    })
                elif isinstance(item, str):
                    normalized_slides.append({
                        "id": str(uuid.uuid4()),
                        "title": item,
                        "number": "",
                        "subtitle": ""
                    })
        elif isinstance(data, dict):
            normalized_slides.append({
                "id": str(uuid.uuid4()),
                "title": str(data.get("title", "")),
                "number": str(data.get("number", "")),
                "subtitle": str(data.get("subtitle", ""))
            })

        return normalized_slides
    except Exception as e:
        st.error(f"AIからのレスポンス解析に失敗しました: {e}\n\n生の応答:\n{raw_response}")
        return None


# --- PowerPoint (.pptx) 生成関数 ---
def create_pptx(slides, theme):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    if theme == "Apple White":
        bg_rgb = RGBColor(255, 255, 255)
        text_rgb = RGBColor(20, 20, 20)
        accent_rgb = RGBColor(0, 102, 204)
        sub_rgb = RGBColor(100, 100, 100)
    elif theme == "Cyber":
        bg_rgb = RGBColor(10, 10, 26)
        text_rgb = RGBColor(240, 240, 255)
        accent_rgb = RGBColor(255, 0, 128)
        sub_rgb = RGBColor(150, 150, 200)
    elif theme == "Minimal Grey":
        bg_rgb = RGBColor(240, 240, 240)
        text_rgb = RGBColor(30, 30, 30)
        accent_rgb = RGBColor(80, 80, 80)
        sub_rgb = RGBColor(120, 120, 120)
    else:
        bg_rgb = RGBColor(0, 0, 0)
        text_rgb = RGBColor(255, 255, 255)
        accent_rgb = RGBColor(0, 229, 255)
        sub_rgb = RGBColor(180, 180, 180)

    for slide_data in slides:
        slide = prs.slides.add_slide(blank_layout)
        
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = bg_rgb

        if slide_data.get("number"):
            txBox = slide.shapes.add_textbox(Inches(1), Inches(1.2), Inches(11.333), Inches(2.0))
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = str(slide_data["number"])
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(72)
            p.font.bold = True
            p.font.color.rgb = accent_rgb

        y_pos = Inches(3.6) if slide_data.get("number") else Inches(2.5)
        txBox2 = slide.shapes.add_textbox(Inches(1), y_pos, Inches(11.333), Inches(1.8))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = slide_data.get("title", "")
        p2.alignment = PP_ALIGN.CENTER
        p2.font.size = Pt(36)
        p2.font.bold = True
        p2.font.color.rgb = text_rgb

        if slide_data.get("subtitle"):
            y_sub_pos = y_pos + Inches(1.8)
            txBox3 = slide.shapes.add_textbox(Inches(1), y_sub_pos, Inches(11.333), Inches(1.2))
            tf3 = txBox3.text_frame
            tf3.word_wrap = True
            p3 = tf3.paragraphs[0]
            p3.text = slide_data.get("subtitle", "")
            p3.alignment = PP_ALIGN.CENTER
            p3.font.size = Pt(20)
            p3.font.color.rgb = sub_rgb

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


# --- メイン画面レイアウト ---
def main():
    st.title("🍏 Minivibe")
    st.caption("Apple風ミニマリズムプレゼン生成＆リアルタイムエディタ")

    with st.sidebar:
        st.header("⚙️ 設定")
        cohere_key = st.text_input("Cohere API Key (推奨)", type="password")
        g_key = st.text_input("Google API Key (Gemini)", type="password")
        groq_key = st.text_input("Groq API Key", type="password")
        
        st.divider()
        theme = st.selectbox(
            "🎨 デザインテーマ",
            ["Apple Black", "Apple White", "Cyber", "Minimal Grey"]
        )

    input_text = st.text_area(
        "プレゼンの原稿や長文アイデアを入力してください",
        height=150,
        placeholder="ここにプレゼンの内容を貼り付けます..."
    )

    if st.button("🚀 AIスライド生成", type="primary"):
        if not (cohere_key or g_key or groq_key):
            st.error("少なくとも1つのAPIキーをサイドバーに入力してください。")
            return
        
        if not input_text.strip():
            st.warning("原稿テキストを入力してください。")
            return

        with st.spinner("Apple風ミニマリズム表現に変換中..."):
            parsed_slides = generate_slides_with_ai(input_text, cohere_key, g_key, groq_key)
            if parsed_slides:
                st.session_state.slides = parsed_slides
                st.success("スライドの生成が完了しました！")
                st.rerun()

    st.divider()

    if st.session_state.slides:
        col_edit, col_preview = st.columns([1, 1])

        # --- 左カラム：編集パネル ---
        with col_edit:
            st.subheader("🎛️ スライド編集パネル")

            if st.button("➕ 新しいスライドを追加"):
                st.session_state.slides.append({
                    "id": str(uuid.uuid4()),
                    "title": "新規スライド",
                    "number": "100%",
                    "subtitle": "補足説明"
                })
                st.rerun()

            # セッション内の要素を走査
            for idx in range(len(st.session_state.slides)):
                slide = st.session_state.slides[idx]
                
                # 型チェック（万が一辞書型でない場合は変換）
                if not isinstance(slide, dict):
                    slide = {"id": str(uuid.uuid4()), "title": str(slide), "number": "", "subtitle": ""}
                    st.session_state.slides[idx] = slide

                if "id" not in slide:
                    slide["id"] = str(uuid.uuid4())

                slide_id = slide["id"]
                slide_title = str(slide.get("title", f"スライド {idx + 1}"))

                with st.accordion(f"スライド {idx + 1}: {slide_title}", expanded=(idx == 0)):
                    new_number = st.text_input("強調数値・キーワード", value=str(slide.get("number", "")), key=f"num_{slide_id}")
                    new_title = st.text_input("メインコピー", value=str(slide.get("title", "")), key=f"title_{slide_id}")
                    new_subtitle = st.text_area("補足テキスト", value=str(slide.get("subtitle", "")), key=f"sub_{slide_id}")

                    # 値の変更を即時反映
                    slide["number"] = new_number
                    slide["title"] = new_title
                    slide["subtitle"] = new_subtitle

                    b_col1, b_col2, b_col3 = st.columns(3)
                    with b_col1:
                        if idx > 0 and st.button("⬆️ 上へ", key=f"up_{slide_id}"):
                            st.session_state.slides[idx], st.session_state.slides[idx-1] = st.session_state.slides[idx-1], st.session_state.slides[idx]
                            st.rerun()
                    with b_col2:
                        if idx < len(st.session_state.slides) - 1 and st.button("⬇️ 下へ", key=f"down_{slide_id}"):
                            st.session_state.slides[idx], st.session_state.slides[idx+1] = st.session_state.slides[idx+1], st.session_state.slides[idx]
                            st.rerun()
                    with b_col3:
                        if st.button("🗑️ 削除", key=f"del_{slide_id}"):
                            st.session_state.slides.pop(idx)
                            st.rerun()

        # --- 右カラム：プレビュー＆ダウンロード ---
        with col_preview:
            st.subheader("👁️ リアルタイムプレビュー")

            bg_css = "#000000"
            text_css = "#ffffff"
            accent_css = "#00e5ff"
            sub_css = "#b0b0b0"

            if theme == "Apple White":
                bg_css = "#ffffff"
                text_css = "#111111"
                accent_css = "#0066cc"
                sub_css = "#666666"
            elif theme == "Cyber":
                bg_css = "#0a0a1a"
                text_css = "#f0f0ff"
                accent_css = "#ff0080"
                sub_css = "#9696c8"
            elif theme == "Minimal Grey":
                bg_css = "#f0f0f0"
                text_css = "#1e1e1e"
                accent_css = "#505050"
                sub_css = "#787878"

            for idx, slide in enumerate(st.session_state.slides):
                if not isinstance(slide, dict):
                    continue

                num_html = f'<div style="font-size: 48px; font-weight: bold; color: {accent_css}; margin-bottom: 8px;">{slide.get("number", "")}</div>' if slide.get("number") else ""
                sub_html = f'<div style="font-size: 14px; color: {sub_css}; margin-top: 8px;">{slide.get("subtitle", "")}</div>' if slide.get("subtitle") else ""

                st.markdown(
                    f"""
                    <div style="background-color: {bg_css}; color: {text_css}; padding: 32px; border-radius: 16px; margin-bottom: 20px; text-align: center; border: 1px solid #333;">
                        <div style="font-size: 12px; color: {sub_css}; margin-bottom: 12px;">SLIDE {idx + 1}</div>
                        {num_html}
                        <div style="font-size: 22px; font-weight: bold; color: {text_css};">{slide.get("title", "")}</div>
                        {sub_html}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            st.divider()
            
            pptx_data = create_pptx(st.session_state.slides, theme)
            st.download_button(
                label="📥 PowerPoint (.pptx) をダウンロード",
                data=pptx_data,
                file_name="Minivibe_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary"
            )

if __name__ == "__main__":
    main()
