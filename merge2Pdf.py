import os
from idlelib.replace import replace
from sys import argv
from tkinter import Tk, filedialog
from PyPDF2 import PdfWriter
from PyPDF2.generic import AnnotationBuilder
from PIL import Image
from docx2pdf import convert
from subprocess import run

debug_mode = False
open_file_on_finish = True
open_path_on_finish = True
supported_types = [("All files", "*.*"),("PDF", ".pdf"), ("JPG", ".jpg"), ("JPEG", ".jpeg"), ("PNG", ".png"), ("DOCX", ".docx"),
                   ("TXT", ".txt")]  # .pptx?

output = PdfWriter()

def get_paths():
    while True:
        if len(argv) <= 1:
            root = Tk()
            root.withdraw()
            file_path = filedialog.askopenfilenames(
                title="Choose files",
                filetypes= supported_types

            )
            if file_path:
                return file_path
            else:
                ch = input("No files, repeat?(y/n): ")
                if ch == 'y' or ch == 'Y':
                    continue
                else:
                    exit()
        else:
            return argv[1:]


def print_order(filelist):
    print("Selected files: ")
    for i in range(len(filelist)):
        print(str(i + 1) + '.', os.path.basename(filelist[i]))


def merge(filelist):
    for path in filelist:
        print('Merging', f'"{os.path.basename(path)}"' + '...')
        ext = file_ext(path)
        if ext == ".pdf":
            output.append(path)
        elif ext in (".jpg", ".jpeg", ".png"):
            image_handler(path)
        elif ext == ".docx":
            docx_handler(path)
        elif ext == ".txt":
            txt_handler(path)
    print("Merged!\n")


def image_handler(path):
    w, h = Image.open(path).size
    output.addBlankPage(width=w, height=h)
    with open(path, "rb") as f:
        output.add_attachment(path, f.read())


def docx_handler(path):
    buff_path = os.path.dirname(path)+r"\Merge2PdfBuff.pdf"
    with open(buff_path,"wb") as f:
        print("Created buffer file")
    convert(path, buff_path)
    output.append(buff_path)
    if os.path.exists(buff_path):
        os.remove(buff_path)
        print("Buffer file deleted")


def txt_handler(path):
    output.addBlankPage(width=595, height=842)
    with open(path, "r") as f:
        annotation = AnnotationBuilder.free_text(
            f.read(),
            rect=(25, 25, 595 - 25, 842 - 25),
            font="Times New Roman",
            border_color="ffffff",
        )
        output.add_annotation(len(output.pages), annotation)


def welcome_message():
    print("Merge2Pdf")
    print("Supported types:", [i[1] for i in supported_types[1:]])
    print()


def file_ext(path):
    return path[len(path) - path[::-1].index('.') - 1:]


def create_pdf(file_reference):
    name = os.path.basename(file_reference)
    # output_name = name.replace(file_ext(name), '') + ".pdf"
    output_name = "output.pdf"
    output_path = file_reference.replace(name, output_name)
    with open(output_path, "wb") as f:
        output.write(f)
    return output_path


def main():

    if not debug_mode:
        try:
            welcome_message()
            filelist = get_paths()
            print_order(filelist)
            merge(filelist)
            output_path = create_pdf(filelist[0])

        except Exception as e:
            print("An Error occured. Error message:")
            print(e)

        else:
            print("File created at", output_path)
            if open_path_on_finish:
                run(['explorer', '/select,', output_path])
            if open_file_on_finish:
                os.startfile(output_path)

        finally:
            print("\nPress any key to exit.")
            input()
            exit()
    else:
        welcome_message()
        filelist = get_paths()
        print_order(filelist)
        merge(filelist)
        output_path = create_pdf(filelist[0])

if __name__ == "__main__":
    main()
