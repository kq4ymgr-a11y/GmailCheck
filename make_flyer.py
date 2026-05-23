from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as ns
from lxml import etree

# A4横: 297mm x 210mm
prs = Presentation()
prs.slide_width  = Inches(11.69)  # 297mm
prs.slide_height = Inches(8.27)   # 210mm

slide_layout = prs.slide_layouts[6]  # blank
slide = prs.slides.add_slide(slide_layout)

W = prs.slide_width
H = prs.slide_height

def add_rect(slide, left, top, width, height, fill_rgb=None, line_rgb=None, line_width_pt=0):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    fill = shape.fill
    if fill_rgb:
        fill.solid()
        fill.fore_color.rgb = RGBColor(*fill_rgb)
    else:
        fill.background()
    line = shape.line
    if line_rgb:
        line.color.rgb = RGBColor(*line_rgb)
        line.width = Pt(line_width_pt)
    else:
        line.fill.background()
    return shape

def add_text(slide, text, left, top, width, height,
             font_size=24, bold=False, color=(0,0,0),
             align=PP_ALIGN.CENTER, font_name="Meiryo"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    run.font.color.rgb = RGBColor(*color)
    return txBox

# --- 背景: 明るいクリーム色 ---
add_rect(slide, 0, 0, W, H, fill_rgb=(255, 253, 230))

# --- 上部の緑帯 ---
banner_h = Inches(1.6)
add_rect(slide, 0, 0, W, banner_h, fill_rgb=(56, 142, 60))

# --- タイトル ---
add_text(slide, "🪲  カブトムシの幼虫、配布中！  🪲",
         Inches(0.3), Inches(0.1), W - Inches(0.6), Inches(1.4),
         font_size=44, bold=True, color=(255, 255, 255),
         font_name="Meiryo")

# --- サブタイトル帯 ---
sub_h = Inches(0.6)
add_rect(slide, 0, banner_h, W, sub_h, fill_rgb=(129, 199, 132))

add_text(slide, "ご自由にお持ちください！  数に限りがあります。お早めに！",
         Inches(0.3), banner_h + Inches(0.05), W - Inches(0.6), sub_h,
         font_size=22, bold=True, color=(27, 94, 32),
         font_name="Meiryo")

# --- 中央エリア ---
content_top = banner_h + sub_h + Inches(0.15)

# 左側: 幼虫イラスト（絵文字テキスト）
emoji_box = add_rect(slide, Inches(0.3), content_top,
                     Inches(3.2), Inches(4.5),
                     fill_rgb=(200, 230, 201), line_rgb=(56,142,60), line_width_pt=2)

add_text(slide, "🥚\n🪱\n🌱",
         Inches(0.3), content_top + Inches(0.2),
         Inches(3.2), Inches(4.0),
         font_size=64, bold=False, color=(0, 0, 0),
         font_name="Meiryo")

add_text(slide, "かわいい幼虫たち",
         Inches(0.3), content_top + Inches(3.4),
         Inches(3.2), Inches(0.6),
         font_size=16, bold=True, color=(27, 94, 32),
         font_name="Meiryo")

# 右側: 説明テキスト
right_left = Inches(3.8)
right_w = W - right_left - Inches(0.3)

# 説明ボックス
add_rect(slide, right_left, content_top,
         right_w, Inches(4.5),
         fill_rgb=(255, 255, 255), line_rgb=(129,199,132), line_width_pt=2)

details = [
    ("📦  数 量", "1人 2〜3匹まで"),
    ("🌿  状 態", "元気な幼虫（2令〜3令）"),
    ("🪣  容 器", "ご持参いただくか、\n    ご相談ください"),
    ("🪵  エ サ", "腐葉土・マットをお渡しします"),
    ("📅  期 間", "なくなり次第終了"),
]

dy = Inches(0.05)
for i, (label, val) in enumerate(details):
    row_top = content_top + Inches(0.15) + i * Inches(0.82)
    # ラベル背景
    add_rect(slide, right_left + Inches(0.15), row_top + Inches(0.05),
             Inches(1.9), Inches(0.6),
             fill_rgb=(200, 230, 201))
    add_text(slide, label,
             right_left + Inches(0.15), row_top + Inches(0.05),
             Inches(1.9), Inches(0.6),
             font_size=16, bold=True, color=(27, 94, 32),
             font_name="Meiryo")
    add_text(slide, val,
             right_left + Inches(2.1), row_top,
             right_w - Inches(2.3), Inches(0.8),
             font_size=17, bold=False, color=(33, 33, 33),
             align=PP_ALIGN.LEFT, font_name="Meiryo")

# --- 下部: 連絡先 ---
footer_top = content_top + Inches(4.65)
footer_h = H - footer_top
add_rect(slide, 0, footer_top, W, footer_h, fill_rgb=(46, 125, 50))

add_text(slide,
         "✉  お問い合わせ・受け取り場所など、お気軽にご連絡ください！",
         Inches(0.3), footer_top + Inches(0.05),
         W - Inches(0.6), footer_h - Inches(0.1),
         font_size=20, bold=True, color=(255, 255, 255),
         font_name="Meiryo")

# --- 右下: 注意書き ---
add_text(slide,
         "※ カブトムシの成虫になったら、自然に返してあげてください 🌳",
         Inches(0.3), footer_top - Inches(0.45),
         W - Inches(0.6), Inches(0.4),
         font_size=13, bold=False, color=(85, 85, 85),
         align=PP_ALIGN.RIGHT, font_name="Meiryo")

out_path = "/home/user/GmailCheck/kabuto_flyer.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
