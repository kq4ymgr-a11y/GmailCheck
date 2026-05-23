from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm
import pptx.oxml.ns as nsmap
from lxml import etree

# A4 landscape: 297mm x 210mm
prs = Presentation()
prs.slide_width = Cm(29.7)
prs.slide_height = Cm(21.0)

slide_layout = prs.slide_layouts[6]  # blank
slide = prs.slides.add_slide(slide_layout)

W = prs.slide_width
H = prs.slide_height

def add_rect(slide, x, y, w, h, fill_color=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE
    fill = shape.fill
    if fill_color:
        fill.solid()
        fill.fore_color.rgb = RGBColor(*fill_color)
    else:
        fill.background()
    line = shape.line
    if line_color:
        line.color.rgb = RGBColor(*line_color)
        if line_width:
            line.width = line_width
    else:
        line.fill.background()
    return shape

def add_textbox(slide, x, y, w, h, text, font_size, bold=False, color=(0,0,0),
                align=PP_ALIGN.CENTER, font_name="BIZ UDPGothic", line_spacing=None):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True
    from pptx.util import Pt
    from pptx.oxml.ns import qn
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)
    run.font.name = font_name
    if line_spacing:
        from pptx.util import Pt as Pt2
        from pptx.oxml.ns import qn as qn2
        pPr = p._pPr
        if pPr is None:
            pPr = p._p.get_or_add_pPr()
        lnSpc = etree.SubElement(pPr, qn2('a:lnSpc'))
        spcPts = etree.SubElement(lnSpc, qn2('a:spcPts'))
        spcPts.set('val', str(int(line_spacing * 100)))
    return txBox

# ── 背景：明るい黄緑 ──
bg = add_rect(slide, 0, 0, W, H, fill_color=(230, 255, 200))

# ── 上部：濃い緑帯 ──
add_rect(slide, 0, 0, W, Cm(3.5), fill_color=(34, 120, 20))

# 上部帯のタイトル
tb = add_textbox(slide, Cm(0.5), Cm(0.2), W - Cm(1), Cm(3.0),
                 "🪲　カブトムシの幼虫　無料配布中　🪲",
                 font_size=32, bold=True, color=(255, 255, 255))

# ── メインキャッチコピー（上帯直下、全幅） ──
add_textbox(slide, Cm(1), Cm(3.8), W - Cm(2), Cm(2.5),
            "ご自由にお持ちください！",
            font_size=44, bold=True, color=(34, 100, 10))

# ── 左カラム：大きな虫の絵文字（y=6.6〜15.5） ──
add_textbox(slide, Cm(1.5), Cm(6.6), Cm(9), Cm(8.5),
            "🐛",
            font_size=100, bold=False, color=(0, 0, 0))

# ── 右カラム：説明テキスト（y=6.6〜15.5） ──
desc = (
    "🌿  今年の夏、カブトムシを育ててみませんか？\n\n"
    "🪱  幼虫の状態でお渡しします（腐葉土入りカップ）\n\n"
    "🌳  お子様の自由研究や夏の思い出づくりに！\n\n"
    "📦  数に限りがありますのでお早めにどうぞ"
)
from pptx.util import Pt
from pptx.dml.color import RGBColor as RGB

tb2 = slide.shapes.add_textbox(Cm(11.5), Cm(6.6), Cm(17.0), Cm(9.0))
tf2 = tb2.text_frame
tf2.word_wrap = True
lines = desc.split("\n")
first = True
for line in lines:
    if first:
        p = tf2.paragraphs[0]
        first = False
    else:
        p = tf2.add_paragraph()
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = line
    run.font.size = Pt(17)
    run.font.bold = False
    run.font.color.rgb = RGB(30, 80, 10)
    run.font.name = "BIZ UDPGothic"

# ── 下部：緑帯 ──
add_rect(slide, 0, Cm(16.5), W, Cm(4.5), fill_color=(34, 120, 20))

add_textbox(slide, Cm(1), Cm(17.2), W - Cm(2), Cm(3.0),
            "岩瀬スポーツ公園管理事務所",
            font_size=26, bold=True, color=(255, 255, 255))

# ── 装飾：枠線 ──
border = add_rect(slide, Cm(0.4), Cm(3.5), W - Cm(0.8), Cm(12.6),
                  line_color=(34, 120, 20), line_width=Pt(2))
border.fill.background()

out = "/home/user/GmailCheck/カブトムシ幼虫配布チラシ.pptx"
prs.save(out)
print(f"Saved: {out}")
