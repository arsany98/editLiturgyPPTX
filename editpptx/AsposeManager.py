from aspose.slides import FontsLoader, export, FontData, Table, ShapeType
from aspose.slides import Presentation as slidesPresentation
from fontTools import ttLib
import os
from pptx import Presentation


class AsposeManager:

    def remove_water_mark(self, file):
        ppt = Presentation(file)
        for slide in ppt.slides:
            for shape in slide.shapes:
                if (
                    shape.has_text_frame
                    and shape.text_frame.text.lower().find("aspose") != -1
                ):
                    slide.shapes.element.remove(shape.element)
        ppt.save(file)

    def get_font_files(self, dir_path):
        result = []
        for root, dirs, files in os.walk(dir_path):
            for f in files:
                file = os.path.join(root, f)

                if file.endswith(".ttf"):
                    font = ttLib.TTFont(file)
                    result.append(FontData(font["name"].getDebugName(1)))
        return result

    def embed_fonts(self, file, font_dir):
        presentation = slidesPresentation(file)
        allFonts = self.get_font_files(font_dir)
        embeddedFonts = presentation.fonts_manager.get_embedded_fonts()
        embeddedFonts = [x.font_name for x in embeddedFonts]

        FontsLoader.load_external_fonts([font_dir])
        for font in allFonts:
            if font.font_name not in embeddedFonts:
                try:
                    presentation.fonts_manager.add_embedded_font(
                        font, export.EmbedFontCharacters.ALL
                    )
                except Exception as e:
                    print(f"{font} {e}")
            else:
                print(f"{font} Font is already embedded.")
        presentation.save(file, export.SaveFormat.PPTX)

    def get_master_line(self, slide_master):
        shapes = slide_master.shapes
        for shape in shapes:
            if shape.shape_type == ShapeType.LINE:
                return shape
        return None

    def move_table_to_master_line(self, file):
        presentation = slidesPresentation(file)
        line = self.get_master_line(presentation.masters[0])
        for slide in presentation.slides:
            tables = 0
            for shape in slide.shapes:
                if type(shape) == Table:
                    tables += 1
            for shape in slide.shapes:
                if line is not None and tables == 1 and type(shape) == Table:
                    shape.y = line.y - shape.rows[0].height
        presentation.save(file, export.SaveFormat.PPTX)
