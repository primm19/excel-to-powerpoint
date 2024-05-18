from tkinter import font
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from PIL import Image
import os

Image.MAX_IMAGE_PIXELS = None

workbook = load_workbook('NMPA Contest Winners 2023 For Power Point.xlsx')
worksheet = workbook.active

range = worksheet['A4':'E269']

presentation = Presentation()

def set_text_and_color(paragraph, text, font_size, font_color):
    p = paragraph
    p.text = text
    font = p.font
    font.size = font_size
    font.color.rgb = font_color

def MakeIntroSlide():
    title_slide_layout = presentation.slide_layouts[0]
    slide = presentation.slides.add_slide(title_slide_layout)

    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)

    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "New Mexico Press Association Awards Presentation"
    title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 0, 0)

    subtitle.text = "2023"
    subtitle.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 0)

def MakeSlides(range):
    i = 1 # Loop variable
    img_num = 2
    for cell in range:
        for x in cell:
            if i == 1:
                slide_register = presentation.slide_layouts[5]
                slide = presentation.slides.add_slide(slide_register)
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(0, 0, 0)
                title = slide.shapes.title

                img_path = 'Contest Winners Art for Powerpont/Imgs/Renames/' + str(img_num) + '.jpg'
                try:
                    slide.shapes.add_picture(img_path, Inches(0.5), Inches(1), Inches(4), Inches(6))
                except (FileNotFoundError, ValueError):
                    pass
                img_num += 1

                title.text = x.value
                title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 0, 0)
                title.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

                i += 1
            elif i == 2:
                left = Inches(6.33)
                top = Inches(2)
                width = height = Inches(3.5)

                txBox = slide.shapes.add_textbox(left, top, width, height)
                tf = txBox.text_frame
                tf.word_wrap = True

                set_text_and_color(tf.add_paragraph(), x.value, Pt(24), RGBColor(255, 255, 0))

                tf.add_paragraph()

                i += 1
            elif i == 3:

                set_text_and_color(tf.add_paragraph(), x.value, Pt(28), RGBColor(255, 255, 0))

                tf.add_paragraph()
                
                i += 1
            elif i == 4:

                set_text_and_color(tf.add_paragraph(), x.value, Pt(28), RGBColor(255, 255, 0))

                tf.add_paragraph()

                i += 1
            elif i >= 5:

                set_text_and_color(tf.add_paragraph(), x.value, Pt(28), RGBColor(255, 255, 0))

                tf.add_paragraph()

                i = 1

MakeIntroSlide()

MakeSlides(range)

presentation.save('test.pptx')