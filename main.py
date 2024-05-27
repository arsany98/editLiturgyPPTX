from tkinter import *
from tkinter import ttk, filedialog
import os
import asyncio
from editpptx.PythonPPTXManager import PythonPPTXManager
from editpptx.AsposeManager import AsposeManager
from fontconverter.FontConverter import FontConverter
from editpptx.TextReplacer import TextReplacer
from editpptx.TextBoxesFormatter import TextBoxesFormatter
from pptx.util import Cm


async def get_pptx_files(dir_path):
    result = []
    for root, dirs, files in os.walk(dir_path):
        for f in files:
            file = os.path.join(root, f)
            if file.endswith(".pptx"):
                result.append(file)
    return result


aspose_manager = AsposeManager()


def convert_fonts_ui(root, files):
    frm = ttk.Frame(root, padding=10)
    frm.grid()

    options = [
        "CS New Athanasius",
        "Avva_Shenouda",
        "Abraam",
        "Unicode",
    ]

    src = StringVar()
    dest = StringVar()

    src.set(options[0])
    dest.set(options[0])

    ttk.Label(frm, text="From:").grid(row=0, column=0)
    ttk.OptionMenu(frm, src, None, *options).grid(row=0, column=1)
    ttk.Label(frm, text="To:").grid(row=0, column=2)
    ttk.OptionMenu(frm, dest, None, *options).grid(row=0, column=3)

    def edit_all_files():
        for i, file in enumerate(files.get()):
            font_converter = FontConverter()
            font_converter.convert_all_text(file, src.get(), dest.get())
            font_converter.change_font(file, src.get(), dest.get())
            aspose_manager.remove_water_mark(file)
            print(str(i + 1) + ". " + file)

    ttk.Button(frm, text="Apply", command=edit_all_files).grid(row=2, column=2)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(column=3, row=2)
    root.mainloop()


def embedded_font_ui(root, files):
    font_dir = StringVar()

    def choose_fonts_dir():
        path = filedialog.askdirectory(title="Choose font files directory:")
        if path is not None and path != "":
            font_dir.set(path)
            font_dir_label.set(path)

    frm = ttk.Frame(root, padding=10)
    frm.grid()
    ttk.Label(frm, text="Fonts Directory: ").grid(row=1, column=0)
    ttk.Button(frm, text="Choose Dir...", command=choose_fonts_dir).grid(
        row=1, column=1
    )
    font_dir_label = StringVar(value="No directory chosen.")
    ttk.Label(frm, textvariable=font_dir_label).grid(row=1, column=3)

    def edit_all_files():
        for i, file in enumerate(files.get()):
            if font_dir.get() != "":
                aspose_manager.embed_fonts(file, font_dir.get())
            aspose_manager.remove_water_mark(file)
            print(str(i + 1) + ". " + file)

    ttk.Button(frm, text="Apply", command=edit_all_files).grid(row=3, column=2)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(row=3, column=3)
    root.mainloop()


def replace_text_ui(root, files):
    frm = ttk.Frame(root, padding=10)
    frm.grid()
    find = StringVar()
    replace = StringVar()
    ttk.Label(frm, text="Find:").grid(row=0, column=0)
    ttk.Entry(frm, textvariable=find).grid(row=0, column=1)
    ttk.Label(frm, text="Replace:").grid(row=0, column=2)
    ttk.Entry(frm, textvariable=replace).grid(row=0, column=3)

    def edit_all_files():
        for i, file in enumerate(files.get()):
            replacer = TextReplacer()
            replacer.edit_ppt(file, find.get(), replace.get())
            print(str(i + 1) + ". " + file)

    ttk.Button(frm, text="Apply", command=edit_all_files).grid(row=2, column=2)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(column=3, row=2)
    root.mainloop()


def format_textbox_ui(root, files):
    frm = ttk.Frame(root, padding=10)
    frm.grid()
    ttk.Label(frm, text="line width:").grid(row=1, column=0)
    line_width = DoubleVar()
    ttk.Entry(frm, textvariable=line_width, width=5).grid(row=1, column=1)

    def edit_all_files():
        for i, file in enumerate(files.get()):
            formatter = TextBoxesFormatter(line_width.get())
            formatter.edit_ppt(file)
            print(str(i + 1) + ". " + file)

    ttk.Button(frm, text="Apply", command=edit_all_files).grid(row=2, column=2)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(column=3, row=2)
    root.mainloop()


def reformat_slides_ui(root, files):
    frm = ttk.Frame(root, padding=10)
    frm.grid()
    width = DoubleVar()
    height = DoubleVar()
    options = ["Custom", "Standard(4:3)", "Widescreen(16:9)"]
    aspect_ratio = StringVar()
    aspect_ratio.set(options[0])
    ttk.Label(frm, text="Aspect Ratio:").grid(row=0, column=0)

    def on_select(selected):
        if selected == options[1]:
            width.set(25.4)
            height.set(19.05)
        elif selected == options[2]:
            width.set(33.867)
            height.set(19.05)

    ttk.OptionMenu(
        frm, aspect_ratio, aspect_ratio.get(), *options, command=on_select
    ).grid(row=0, column=1)
    ttk.Label(frm, text="New width:").grid(row=1, column=0)
    ttk.Entry(frm, textvariable=width, width=5).grid(row=1, column=1)
    ttk.Label(frm, text="Cm").grid(row=1, column=1, sticky=E)
    ttk.Label(frm, text="New height:").grid(row=1, column=2)
    ttk.Entry(frm, textvariable=height, width=5).grid(row=1, column=3)
    ttk.Label(frm, text="Cm").grid(row=1, column=3, sticky=E)

    ttk.Label(frm, text="Arabic Font Size Increase:").grid(row=2, column=0)
    font_size_increase = IntVar()
    ttk.Entry(frm, textvariable=font_size_increase, width=4).grid(row=2, column=1)
    ttk.Label(frm, text="Pt").grid(row=2, column=1, sticky=E)

    ttk.Label(frm, text="Coptic Font Size Increase:").grid(row=2, column=2)
    copt_font_size_increase = IntVar()
    ttk.Entry(frm, textvariable=copt_font_size_increase, width=4).grid(row=2, column=3)
    ttk.Label(frm, text="Pt").grid(row=2, column=3, sticky=E)

    exclude_first_slide = BooleanVar()
    move_table_to_master_line = BooleanVar()

    def edit_all_files():
        editor = PythonPPTXManager(
            Cm(width.get()),
            Cm(height.get()),
            font_size_increase.get(),
            copt_font_size_increase.get(),
            exclude_first_slide.get(),
        )
        for i, file in enumerate(files.get()):
            editor.edit_ppt(file)
            if move_table_to_master_line:
                aspose_manager.move_table_to_master_line(file)
                aspose_manager.remove_water_mark(file)
            print(str(i + 1) + ". " + file)

    ttk.Checkbutton(frm, text="Exclude First Slide", variable=exclude_first_slide).grid(
        row=3, column=0
    )
    ttk.Checkbutton(
        frm, text="Move Tables to master line", variable=move_table_to_master_line
    ).grid(row=3, column=1)
    ttk.Button(frm, text="Apply", command=edit_all_files).grid(row=4, column=2)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(row=4, column=3)


def GUI():
    root = Tk()
    root.title("Edit PPTX")
    root.resizable(False, False)
    files = Variable(value=[])

    def choose_dir():
        path = filedialog.askdirectory(title="Choose pptx files directory:")
        if path is not None and path != "":
            files.set(asyncio.run(get_pptx_files(path)))
            dir_label.set(f"Found {len(files.get())} pptx files")

    def choose_files():
        paths = filedialog.askopenfilenames(
            title="Choose pptx files:", filetypes=[("power point", "pptx")]
        )
        if paths is not None and paths != []:
            files.set(paths)
            dir_label.set(f"Found {len(files.get())} pptx files")

    frm = ttk.Frame(root, padding=10)
    frm.grid()
    ttk.Label(frm, text="Choose pptx files: ").grid(row=0, column=0)
    ttk.Button(frm, text="Choose Dir...", command=choose_dir).grid(row=0, column=1)
    ttk.Button(frm, text="Choose Files...", command=choose_files).grid(row=0, column=2)
    dir_label = StringVar(value="No directory chosen.")
    ttk.Label(frm, textvariable=dir_label).grid(row=0, column=3)

    def open_reformat_slides_ui():
        window = Toplevel(root)
        window.grab_set()
        window.title("Reformat Slides")
        window.resizable(False, False)
        reformat_slides_ui(window, files)

    def open_embedded_fonts_ui():
        window = Toplevel(root)
        window.grab_set()
        window.title("Embed Fonts to PPTX")
        window.resizable(False, False)
        embedded_font_ui(window, files)

    def open_convert_fonts_ui():
        window = Toplevel(root)
        window.grab_set()
        window.title("Convert Fonts")
        window.resizable(False, False)
        convert_fonts_ui(window, files)

    def open_replace_text_ui():
        window = Toplevel(root)
        window.grab_set()
        window.title("Replace Text")
        window.resizable(False, False)
        replace_text_ui(window, files)

    def open_format_textbox_ui():
        window = Toplevel(root)
        window.grab_set()
        window.title("Format Text Boxes")
        window.resizable(False, False)
        format_textbox_ui(window, files)

    ttk.Button(
        frm, text="Reformat Slides", command=open_reformat_slides_ui, width=64
    ).grid(row=1, columnspan=4)
    ttk.Button(
        frm, text="Embedded Fonts", command=open_embedded_fonts_ui, width=64
    ).grid(row=2, columnspan=4)
    ttk.Button(frm, text="Convert Fonts", command=open_convert_fonts_ui, width=64).grid(
        row=3, columnspan=4
    )
    ttk.Button(frm, text="Replace Text", command=open_replace_text_ui, width=64).grid(
        row=4, columnspan=4
    )
    ttk.Button(
        frm, text="Format Text Boxes", command=open_format_textbox_ui, width=64
    ).grid(row=5, columnspan=4)
    ttk.Button(frm, text="Quit", command=root.destroy, width=64).grid(
        row=6, columnspan=4
    )
    root.mainloop()


if __name__ == "__main__":
    GUI()
