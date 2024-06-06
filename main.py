from tkinter import *
from tkinter import ttk, filedialog, messagebox
import os
import asyncio
from editpptx.PythonPPTXManager import PythonPPTXManager
from editpptx.AsposeManager import AsposeManager
from fontconverter.FontConverter import FontConverter
from editpptx.TextReplacer import TextReplacer
from editpptx.TextBoxesFormatter import TextBoxesFormatter


async def get_pptx_files(dir_path):
    result = []
    for root, dirs, files in os.walk(dir_path):
        for f in files:
            file = os.path.join(root, f)
            if file.endswith(".pptx"):
                result.append(file)
    return result


aspose_manager = AsposeManager()


def validate(value, isInt):
    if value != "":
        try:
            if isInt:
                return int(value)
            else:
                return float(value)
        except Exception as e:
            raise e
    return None


def show_progress_bar_ui(root, title, progress, status):
    window = Toplevel(root)
    window.grab_set()
    window.title(title)
    window.resizable(False, False)
    progress_frm = ttk.Frame(window, padding=10)
    progress_frm.grid()
    ttk.Label(window, textvariable=status).grid(row=0, column=0, sticky=W)
    ttk.Progressbar(window, variable=progress, length=600, mode="determinate").grid(
        row=1, column=0
    )
    return window


def edit_all_files(root, files, edit_file, **kwargs):
    status = StringVar()
    progress = IntVar()
    window = show_progress_bar_ui(root, "Convert Fonts", progress, status)
    log = []
    success_files = 0
    for i, file in enumerate(files.get()):
        try:
            edit_file(file, **kwargs)
            success_files += 1
            status.set(str(i + 1) + ". " + file)
            progress.set((i + 1) / len(files.get()) * 100)
            window.update_idletasks()
        except Exception as e:
            log.append(f"{file} {e}\n")
    window.destroy()
    root.grab_set()

    messagebox.showinfo(
        "Log",
        f"Applied to {success_files} files\n{len(log)} files failed\n" + "".join(log),
    )


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

    def edit_file(file):
        font_converter = FontConverter()
        font_converter.convert_all_text(file, src.get(), dest.get())
        font_converter.change_font(file, src.get(), dest.get())
        aspose_manager.remove_water_mark(file)

    def apply_command():
        edit_all_files(root, files, edit_file)

    ttk.Button(frm, text="Apply", command=apply_command).grid(row=2, column=2)
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

    def edit_file(file):
        if font_dir.get() != "":
            aspose_manager.embed_fonts(file, font_dir.get())
        aspose_manager.remove_water_mark(file)

    def apply_command():
        edit_all_files(root, files, edit_file)

    ttk.Button(frm, text="Apply", command=apply_command).grid(row=3, column=2)
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

    def edit_file(file):
        replacer = TextReplacer()
        replacer.edit_ppt(file, find.get(), replace.get())

    def apply_command():
        edit_all_files(root, files, edit_file)

    ttk.Button(frm, text="Apply", command=apply_command).grid(row=2, column=2)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(column=3, row=2)
    root.mainloop()


def reformat_slides_ui(root, files):
    frm = ttk.Frame(root, padding=10)
    frm.grid()
    width = StringVar()
    height = StringVar()
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
    ttk.Entry(frm, textvariable=width, width=6).grid(row=1, column=1)
    ttk.Label(frm, text="Cm").grid(row=1, column=1, sticky=E)
    ttk.Label(frm, text="New height:").grid(row=1, column=2)
    ttk.Entry(frm, textvariable=height, width=6).grid(row=1, column=3, padx=30)
    ttk.Label(frm, text="Cm").grid(row=1, column=3, sticky=E)

    ttk.Label(frm, text="Arabic Font Size Increase:").grid(row=2, column=0)
    font_size_increase = StringVar()
    ttk.Entry(frm, textvariable=font_size_increase, width=6).grid(row=2, column=1)
    ttk.Label(frm, text="Pt").grid(row=2, column=1, sticky=E)

    ttk.Label(frm, text="Coptic Font Size Increase:").grid(row=2, column=2)
    copt_font_size_increase = StringVar()
    ttk.Entry(frm, textvariable=copt_font_size_increase, width=6).grid(row=2, column=3)
    ttk.Label(frm, text="Pt").grid(row=2, column=3, sticky=E)

    exclude_outlined = BooleanVar()
    ttk.Checkbutton(frm, text="Exclude Outlined", variable=exclude_outlined).grid(
        row=3, column=0
    )

    class Margin:
        def __init__(self):
            self.left = StringVar()
            self.top = StringVar()
            self.right = StringVar()
            self.bottom = StringVar()

        def validate(self):
            try:
                validated = Margin()
                validated.left = validate(self.left.get(), False)
                validated.top = validate(self.top.get(), False)
                validated.right = validate(self.right.get(), False)
                validated.bottom = validate(self.bottom.get(), False)
                return validated
            except Exception as e:
                raise e

        def UI(self, frm, sRow, title):
            ttk.Label(frm, text=title).grid(row=sRow, column=0, rowspan=2, pady=20)
            ttk.Label(frm, text="Left").grid(row=sRow, column=1, sticky=W)
            ttk.Entry(frm, textvariable=self.left, width=6).grid(
                row=sRow, column=1, padx=50
            )
            ttk.Label(frm, text="Cm").grid(row=sRow, column=1, sticky=E)
            ttk.Label(frm, text="Top").grid(row=sRow, column=2, sticky=W)
            ttk.Entry(frm, textvariable=self.top, width=6).grid(
                row=sRow, column=2, padx=50
            )
            ttk.Label(frm, text="Cm").grid(row=sRow, column=2, sticky=E)
            ttk.Label(frm, text="Bottom").grid(row=sRow + 1, column=1, sticky=W)
            ttk.Entry(frm, textvariable=self.bottom, width=6).grid(
                row=sRow + 1, column=1, padx=50
            )
            ttk.Label(frm, text="Cm").grid(row=sRow + 1, column=1, sticky=E)
            ttk.Label(frm, text="Right").grid(row=sRow + 1, column=2, sticky=W)
            ttk.Entry(frm, textvariable=self.right, width=6).grid(
                row=sRow + 1, column=2, padx=50
            )
            ttk.Label(frm, text="Cm").grid(row=sRow + 1, column=2, sticky=E)

    shape_margin = Margin()
    shape_margin.UI(frm, 4, "Shape Margin:")

    table_margin = Margin()
    table_margin.UI(frm, 6, "Table Margin:")

    ttk.Label(frm, text="Outline Width:").grid(row=8, column=0)
    line_width = StringVar()
    ttk.Entry(frm, textvariable=line_width, width=6).grid(row=8, column=1)
    ttk.Label(frm, text="Pt").grid(row=8, column=1, sticky=E)

    ttk.Label(frm, text="Table Line Position:").grid(row=9, column=0)
    table_position = StringVar()
    ttk.Entry(frm, textvariable=table_position, width=6).grid(row=9, column=1)
    ttk.Label(frm, text="Cm").grid(row=9, column=1, sticky=E)

    exclude_first_slide = BooleanVar()
    ttk.Checkbutton(frm, text="Exclude First Slide", variable=exclude_first_slide).grid(
        row=10, column=0
    )

    def edit_file(file, editor, tp):
        editor.edit_ppt(file)
        if tp is not None:
            aspose_manager.move_table_to_position(file, tp)
            aspose_manager.remove_water_mark(file)

    def apply_command():
        try:
            editor = PythonPPTXManager(
                validate(width.get(), False),
                validate(height.get(), False),
                validate(font_size_increase.get(), True),
                validate(copt_font_size_increase.get(), True),
                exclude_first_slide.get(),
                exclude_outlined.get(),
                shape_margin.validate(),
                table_margin.validate(),
                validate(line_width.get(), True),
            )
            tp = validate(table_position.get(), False)
            edit_all_files(root, files, edit_file, editor=editor, tp=tp)
        except Exception as e:
            messagebox.showerror("Invalid Input", e)

    ttk.Button(frm, text="Apply", command=apply_command).grid(row=11, column=3)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(row=11, column=4)
    root.mainloop()


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
            count_slides()

    def choose_files():
        paths = filedialog.askopenfilenames(
            title="Choose pptx files:", filetypes=[("power point", "pptx")]
        )
        if paths is not None and paths != []:
            files.set(paths)
            dir_label.set(f"Found {len(files.get())} pptx files")
            count_slides()

    def count_slides():
        slides_count.set(f"Slides Count: Loading...")
        root.update_idletasks()
        co = 0
        for file in files.get():
            try:
                co += PythonPPTXManager.get_slides_count(file)
            except Exception as e:
                print(e)
        slides_count.set(f"Slides Count: {co}")

    frm = ttk.Frame(root, padding=10)
    frm.grid()
    ttk.Label(frm, text="Choose pptx files: ").grid(row=0, column=0)
    ttk.Button(frm, text="Choose Dir...", command=choose_dir).grid(row=0, column=1)
    ttk.Button(frm, text="Choose Files...", command=choose_files).grid(row=0, column=2)
    dir_label = StringVar(value="No directory chosen.")
    slides_count = StringVar(value="Slides Count: 0")
    ttk.Label(frm, textvariable=dir_label).grid(row=0, column=3)
    ttk.Label(frm, textvariable=slides_count).grid(row=1, columnspan=4)

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

    ttk.Button(
        frm, text="Reformat Slides", command=open_reformat_slides_ui, width=64
    ).grid(row=2, columnspan=4)
    ttk.Button(
        frm, text="Embedded Fonts", command=open_embedded_fonts_ui, width=64
    ).grid(row=3, columnspan=4)
    ttk.Button(frm, text="Convert Fonts", command=open_convert_fonts_ui, width=64).grid(
        row=4, columnspan=4
    )
    ttk.Button(frm, text="Replace Text", command=open_replace_text_ui, width=64).grid(
        row=5, columnspan=4
    )
    ttk.Button(frm, text="Quit", command=root.destroy, width=64).grid(
        row=6, columnspan=4
    )
    root.mainloop()


if __name__ == "__main__":
    GUI()
