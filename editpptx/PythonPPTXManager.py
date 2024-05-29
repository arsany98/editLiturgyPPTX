from pptx.util import Cm, Emu, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL_TYPE
from pptx import Presentation
import re


class PythonPPTXManager:
    def __init__(
        self,
        new_width,
        new_height,
        font_size_increase,
        copt_font_size_increase,
        exclude_first_slide,
        exclude_outlined,
        shape_margin,
        table_margin,
    ):
        self.new_width = new_width
        self.new_height = new_height
        self.font_size_increase = font_size_increase
        self.copt_font_size_increase = copt_font_size_increase
        self.exclude_first_slide = exclude_first_slide
        self.exclude_outlined = exclude_outlined
        self.shape_margin = shape_margin
        self.table_margin = table_margin

    def change_shape_width(self, shape, exclude):
        ratio = self.new_width / self.old_width
        diff = self.new_width - self.old_width
        w = shape.width
        h = shape.height
        l = shape.left
        t = shape.top
        if exclude:
            shape.left = Emu(l + (diff / 2))
            return
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            shape.left = Emu((l * ratio) if l > 0 else l)
        elif shape.is_placeholder:
            shape.width = Emu(w * ratio)
            shape.left = Emu((l * ratio) if l > 0 else l)
            shape.height = Emu(h)
            shape.top = Emu(t)
        else:
            shape.width = Emu(w * ratio)
            shape.left = Emu((l * ratio) if l > 0 else l)
        if shape.has_table:
            for column in shape.table.columns:
                column.width = Emu(column.width * ratio)

    def change_shape_height(self, shape, exclude):
        ratio = self.new_height / self.old_height
        diff = self.new_height - self.old_height
        w = shape.width
        h = shape.height
        l = shape.left
        t = shape.top
        if exclude:
            shape.top = Emu(t + (diff / 2))
            return
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            shape.top = Emu((t * ratio) if t > 0 else t)
        elif shape.is_placeholder:
            shape.height = Emu(h * ratio)
            shape.top = Emu((t * ratio) if t > 0 else t)
            shape.width = Emu(w)
            shape.left = Emu(l)
        else:
            shape.height = Emu(h * ratio)
            shape.top = Emu((t * ratio) if t > 0 else t)
        if shape.has_table:
            for row in shape.table.rows:
                row.height = Emu(row.height * ratio)

    def increase_font_size(self, text_frame, font_increase):
        if self.font_size_increase is not None:
            prgs = text_frame.paragraphs
            for prg in prgs:
                for run in prg.runs:
                    if run.font.size is not None:  # error font less than 0
                        run.font.size += Pt(font_increase)

    def increase_arabic_shape_font_size(self, shape):
        if shape.has_table:
            for cell in shape.table.iter_cells():
                if self.is_arabic(cell.text_frame.text):
                    self.increase_font_size(cell.text_frame, self.font_size_increase)
        if shape.has_text_frame and self.font_size_increase is not None:
            if self.is_arabic(shape.text_frame.text):
                self.increase_font_size(shape.text_frame, self.font_size_increase)

    def increase_coptic_shape_font_size(self, shape):
        if shape.has_table:
            for cell in shape.table.iter_cells():
                if not self.is_arabic(cell.text_frame.text):
                    self.increase_font_size(
                        cell.text_frame, self.copt_font_size_increase
                    )
        if shape.has_text_frame and self.font_size_increase is not None:
            if not self.is_arabic(shape.text_frame.text):
                self.increase_font_size(shape.text_frame, self.copt_font_size_increase)

    def edit_slide(self, slide, exclude=False):
        shapes = slide.shapes
        for shape in shapes:
            if self.new_width is not None and self.new_width != 0:
                self.change_shape_width(shape, exclude)
            if self.new_height is not None and self.new_height != 0:
                self.change_shape_height(shape, exclude)
            if not exclude:
                if (
                    self.exclude_outlined
                    and (
                        shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX
                        or shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                    )
                    and shape.line.fill.type == MSO_FILL_TYPE.SOLID
                ):
                    continue

                if self.font_size_increase is not None and self.font_size_increase != 0:
                    self.increase_arabic_shape_font_size(shape)
                if (
                    self.copt_font_size_increase is not None
                    and self.copt_font_size_increase != 0
                ):
                    self.increase_coptic_shape_font_size(shape)
                self.set_shape_margin(shape)

    def edit_ppt(self, file):
        try:
            ppt = Presentation(file)
            if self.new_width is not None and self.new_width != 0:
                self.old_width = ppt.slide_width
                ppt.slide_width = self.new_width
            if self.new_height is not None and self.new_height != 0:
                self.old_height = ppt.slide_height
                ppt.slide_height = self.new_height

            self.edit_slide(ppt.slide_master)
            for slide_layout in ppt.slide_master.slide_layouts:
                self.edit_slide(slide_layout)

            for idx, slide in enumerate(ppt.slides):
                self.edit_slide(slide, exclude=(self.exclude_first_slide and idx == 0))

            ppt.save(file)
        except Exception as e:
            print(file, "is invalid", e)

    def is_arabic(self, text):
        pattern = re.compile(".*[\\u0600-\\u06FF]")
        match = pattern.match(text)
        return match is not None

    def set_text_margin(self, text_frame, margin):
        text_frame.margin_left = Cm(margin.left.get())
        text_frame.margin_top = Cm(margin.top.get())
        text_frame.margin_right = Cm(margin.right.get())
        text_frame.margin_bottom = Cm(margin.bottom.get())

    def set_shape_margin(self, shape):
        if shape.has_table:
            for cell in shape.table.iter_cells():
                self.set_text_margin(cell, self.table_margin)
        if shape.has_text_frame:
            self.set_text_margin(shape.text_frame, self.shape_margin)
