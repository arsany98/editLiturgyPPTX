from .convert_util import convert
from pptx import Presentation


class FontConverter:
    def convert_font(self, text_frame, src, dest):
        prgs = text_frame.paragraphs
        for prg in prgs:
            for run in prg.runs:
                if run.font.name == src:
                    converted = convert(run.text, src, dest)
                    run.text = converted
                    run.font.name = dest

    def convert_all_text(self, file, src, dest):
        if src == "Unicode":
            src = "Coptic New Athanasius"
        if dest == "Unicode":
            dest = "Coptic New Athanasius"
        ppt = Presentation(file)
        for idx, slide in enumerate(ppt.slides):
            shapes = slide.shapes
            for shape in shapes:
                if shape.has_table:
                    for cell in shape.table.iter_cells():
                        self.convert_font(cell.text_frame, src, dest)
                if shape.has_text_frame:
                    self.convert_font(shape.text_frame, src, dest)
        ppt.save(file)
