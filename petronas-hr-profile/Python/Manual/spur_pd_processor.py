import pandas as pd
import numpy as np
import re
import glob
import shutil
from docx2pdf import convert
import textwrap
from fpdf import FPDF
import os
import sys
import win32com.client as win32
from win32com.client import constants

# import comtypes.client
from tqdm import tqdm

# Get path from command line argument
ABS_PATH = sys.argv[0]


def save_as_docx(path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch("Word.Application")
    doc = word.Documents.Open(path)
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r"\.\w+$", ".pdf", new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=17)
    doc.Close(False)


def main():
    source = ABS_PATH

    for root, dirs, filenames in os.walk(source):
        for f in filenames:
            filename, file_extension = os.path.splitext(f)

            if file_extension.lower() == ".doc":
                file_conv = os.path.join(root, f)
                save_as_docx(file_conv)
                print("%s ==> %sx" % (file_conv, f))


if __name__ == "__main__":
    main()


def pd_processor(position_blob_dir, pd_folder):
    """
    docstring
    """

    doc_list = list(set(glob.glob(pd_folder + "\\*.doc") + glob.glob(pd_folder + "\\*.DOC")))
    docx_list = list(set(glob.glob(pd_folder + "\\*.docx") + glob.glob(pd_folder + "\\*.DOCX")))
    rtf_list = list(set(glob.glob(pd_folder + "\\*.rtf") + glob.glob(pd_folder + "\\*.RTF")))

    word_list = doc_list + docx_list + rtf_list
    word_list = [x for x in word_list if not x.startswith("~")]
    # Convert doc to docx
    print("Converting doc, docx, & rtf to pdf")
    for doc in tqdm(word_list):
        # print(doc)
        save_as_docx(doc)

    # convert rtf to docx and embed all pictures in the final document
    # rtf_list = list(set(glob.glob(pd_folder + "\\*.rtf") + glob.glob(pd_folder + "\\*.RTF")))

    # def ConvertRtfToDocx(rootDir, file, savename):
    #     word = win32.Dispatch("Word.Application")
    #     wdFormatDocumentDefault = 16
    #     wdHeaderFooterPrimary = 1
    #     doc = word.Documents.Open(rootDir + "\\" + file)
    #     for pic in doc.InlineShapes:
    #         pic.LinkFormat.SavePictureWithDocument = True
    #     for hPic in doc.sections(1).headers(wdHeaderFooterPrimary).Range.InlineShapes:
    #         hPic.LinkFormat.SavePictureWithDocument = True
    #     doc.SaveAs(str(rootDir + "\\refman.docx"), FileFormat=wdFormatDocumentDefault)
    #     doc.Close()
    #     word.Quit()

    # for rtf in rtf_list:
    #     filename = re.search(r"[^\\]+$", rtf).group()
    #     ConvertRtfToDocx(
    #         pd_folder,
    #         filename,
    #         filename,
    #     )

    # Convert docx to pdf
    # convert(pd_folder, keep_active=True)

    txt_list = list(set(glob.glob(pd_folder + "\\*.txt") + glob.glob(pd_folder + "\\*.TXT")))
    txt_list = [x for x in txt_list if not x.startswith("~")]

    def text_to_pdf(text, filename):
        a4_width_mm = 260
        pt_to_mm = 0.35
        fontsize_pt = 9
        fontsize_mm = fontsize_pt * pt_to_mm
        margin_bottom_mm = 10
        character_width_mm = 7 * pt_to_mm
        width_text = a4_width_mm / character_width_mm

        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.set_auto_page_break(True, margin=margin_bottom_mm)
        pdf.add_page()
        pdf.set_font(family="Courier", size=fontsize_pt)
        splitted = text.split("\n")

        for line in splitted:
            lines = textwrap.wrap(line, width_text)

            if len(lines) == 0:
                pdf.ln()

            for wrap in lines:
                pdf.cell(0, fontsize_mm, wrap, ln=1)

        pdf.output(filename, "F")

    if len(txt_list) != 0:
        print("Converting txt to pdf")
        for txt in tqdm(txt_list):
            filename = re.search(r"[^\\]+$", txt).group().zfill(8).replace(".txt", "")
            # print(filename)
            with open(txt, "r", errors="ignore", encoding="utf-8") as f2:
                data = f2.read()
                text = (
                    data.replace("ÿþ", "")
                    .replace('" 	', "\x95	")
                    .replace("\x00", "")
                    .replace("\x19 s", "'s")
                    .replace("\u2013", "-")
                    .replace("\u2022", "\x95 ")
                    .replace("\u2019", "'")
                    .replace("\u2018", "'")
                )
                filename_save = os.path.join(pd_folder, filename + ".pdf")
                text_to_pdf(text, filename_save)

    # wdFormatPDF = 17
    # for subdir, dirs, files in os.walk(pd_folder):
    #     for file in files:
    #         if not file.lower().endswith("rtf"):
    #             continue
    #         in_file = os.path.join(subdir, file)
    #         output_file = file.split(".")[0]
    #         out_file = pd_folder + output_file + ".pdf"
    #         word = comtypes.client.CreateObject("Word.Application")

    #         doc = word.Documents.Open(in_file)
    #         doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    #         doc.Close()
    #         word.Quit()

    pdf_list = list(set(glob.glob(pd_folder + "\\*.pdf") + glob.glob(pd_folder + "\\*.PDF")))

    for pdf in pdf_list:
        filename = re.search(r"[^\\]+(?=\.)", pdf).group().zfill(8)
        if filename.startswith("~"):
            continue
        # print(filename)
        shutil.copy2(
            pdf,
            position_blob_dir + "\\PD" + filename + ".pdf",
        )


# pd_processor(
#     r"D:\Project\HR_SPUR\data_migration\FS02_SPUR_migration\Position_SPUR\BlobFiles",
#     r"D:\Project\HR_SPUR\data_migration\FS02_SPUR_migration\data\PD//",
# )
