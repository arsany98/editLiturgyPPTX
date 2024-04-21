from .convert_util import convert
from pptx import Presentation
from aspose.slides import Presentation as slidesPresentation
from aspose.slides import util, export, FontData


class FontConverter:
    def convert_font(self, text_frame, src, dest):
        prgs = text_frame.paragraphs
        for prg in prgs:
            for run in prg.runs:
                if run.font.name == src:
                    converted = convert(run.text, src, dest)
                    run.text = converted

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

    def change_font(self, file, src, dest):
        if src == "Unicode":
            src = "Coptic New Athanasius"
        if dest == "Unicode":
            dest = "Coptic New Athanasius"
        ppt = slidesPresentation(file)
        textFramesPPTX = util.SlideUtil.get_all_text_frames(ppt, True)
        for idx, text_frame in enumerate(textFramesPPTX):
            prgs = text_frame.paragraphs
            for prg in prgs:
                for port in prg.portions:
                    if (
                        port.portion_format.latin_font is not None
                        and port.portion_format.latin_font.font_name == src
                    ):
                        if dest == "Coptic New Athanasius":
                            port.portion_format.latin_font = FontData(dest)
                            port.portion_format.east_asian_font = FontData(dest)
                            port.portion_format.complex_script_font = None
                            port.portion_format.symbol_font = None
                        else:
                            port.portion_format.latin_font = FontData(dest)
                            port.portion_format.east_asian_font = None
                            port.portion_format.complex_script_font = None
                            port.portion_format.symbol_font = None

        ppt.save(file, export.SaveFormat.PPTX)
