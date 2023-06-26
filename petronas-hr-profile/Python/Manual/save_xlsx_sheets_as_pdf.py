import win32com.client
from win32com.client import constants
from glob import glob
import re
import os
import logging.config

log = logging.getLogger(__name__)


def save_xlsx_sheets_as_pdf(
    xlsx_file,
    job_blob_dir,
    job_clob_dir,
    position_blob_dir,
    position_clob_dir,
    skg_name,
    ws_name_list,
):
    excel_object = win32com.client.Dispatch("Excel.Application")
    excel_object.Visible = False

    wb = excel_object.Workbooks.Open(xlsx_file)
    # worksheets_list = [ws.Name for ws in wb.Worksheets]
    # worksheets_list = [ws for ws in ws_name_list if wb.Worksheets(ws).Visible != 0]
    worksheets_list = ws_name_list
    # worksheets_list = [x for x in worksheets_list if isinstance(x, str)]
    print(worksheets_list)

    def pdf_print_formatting(ws_name):
        ws = wb.Worksheets[ws_name]

        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesTall = 1
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.Orientation = 2
        ws.PageSetup.PaperSize = constants.xlPaperA4
        ws.PageSetup.OddAndEvenPagesHeaderFooter = False
        ws.PageSetup.DifferentFirstPageHeaderFooter = False
        ws.PageSetup.ScaleWithDocHeaderFooter = True
        ws.PageSetup.AlignMarginsHeaderFooter = True
        ws.PageSetup.CenterHorizontally = True
        ws.PageSetup.CenterVertically = True
        if skg_name == "SKG001":
            ws.Columns.EntireColumn.Hidden = False
            # ws.Range("C1").Select()
            excel_object.ActiveWindow.FreezePanes = False

    if (skg_name == "SKG001") or (skg_name == "SKG016"):
        for ws_, filename in worksheets_list:
            if wb.Worksheets(ws_).Visible == 0:
                continue
            if re.search("Sheet1", ws_):
                continue
            print(ws_)
            # filename = wb.Worksheets(ws_).Range("A6").Value
            pdf_print_formatting(ws_)
            wb.Worksheets(ws_).Select()
            log.info("[Extract xlsx write up] Generating {}".format(f"{filename}.pdf"))
            wb.ActiveSheet.ExportAsFixedFormat(0, job_blob_dir + "\\" + f"{filename}.pdf")

    # else:
    #     if skg_name == "SKG016":
    #         do_deliver_worksheets_list = [
    #             x for x in worksheets_list if not re.search("^\d|SUMMARY|DISPLAY", x, flags=re.I)
    #         ]
    #         display_worksheets_list = [x for x in worksheets_list if re.search("DISPLAY|\d$", x, flags=re.I)]
    #     elif skg_name == "SKG010":
    #         do_deliver_worksheets_list = [
    #             x for x in worksheets_list if not re.search("DISPLAY|GTI|All", x, flags=re.I)
    #         ]
    #         display_worksheets_list = [x for x in worksheets_list if re.search("DISPLAY", x, flags=re.I)]
    #     else:
    #         do_deliver_worksheets_list = [x for x in worksheets_list if not re.search("DISPLAY", x, flags=re.I)]
    #         display_worksheets_list = [x for x in worksheets_list if re.search("DISPLAY", x, flags=re.I)]

    #     ws_list = []
    #     for index in do_deliver_worksheets_list:
    #         # print(index)
    #         ws = wb.Worksheets[index]
    #         ur_id_name = [x for x in list(ws.Range("A1:J1").Value[0]) if x is not None][0]
    #         filename = ur_id_name.split("|")[0].strip()
    #         ur_name = ur_id_name.split("|")[1].lower().replace("aromatic", "aromatics").replace("and", "&").strip()
    #         ur_name = re.sub("-$", "", ur_name).strip()
    #         if skg_name != "SKG014":
    #             if ur_name.startswith("principal"):
    #                 temp_ur_name = re.sub(", staff", "", ur_name).strip()
    #                 display_sheet = [
    #                     x
    #                     for x in display_worksheets_list
    #                     if ur_name in str(wb.Worksheets[x].Range("B1").Value).replace(", Staff", "").lower()
    #                     or ur_name in str(wb.Worksheets[x].Range("C3").Value).replace(", Staff", "").lower()
    #                     or ur_name in str(wb.Worksheets[x].Range("K10").Value).replace(", Staff", "").lower()
    #                     or ur_name in str(wb.Worksheets[x].Range("A1").Value).replace(", Staff", "").lower()
    #                 ][0]

    #             else:
    #                 display_sheet = [
    #                     x
    #                     for x in display_worksheets_list
    #                     if ur_name in str(wb.Worksheets[x].Range("B1").Value).replace("and", "&").lower()
    #                     or ur_name in str(wb.Worksheets[x].Range("B2").Value).replace("and", "&").lower()
    #                     or ur_name in str(wb.Worksheets[x].Range("C3").Value).replace("and", "&").lower()
    #                     or ur_name in str(wb.Worksheets[x].Range("K10").Value).replace("and", "&").lower()
    #                     or ur_name in str(wb.Worksheets[x].Range("A1").Value).replace("and", "&").lower()
    #                     or ur_name in str(wb.Worksheets[x].Range("A2").Value).replace("and", "&").lower()
    #                 ][0]
    #             print(display_sheet)

    #         else:
    #             display_sheet = display_worksheets_list[0]

    #         ws_list.append([ws.Name, display_sheet])

    #     for ws_pair in ws_list:
    #         print(ws_pair)
    #         ws_1st = wb.Worksheets[ws_pair[0]]
    #         ur_id_name = [x for x in list(ws_1st.Range("A1:J1").Value[0]) if x is not None][0]
    #         filename = re.search(r".+?(?=\|)", ur_id_name).group().strip().replace("SKG", "0").replace("UR ID: ", "")
    #         if f"{filename}.pdf" in os.listdir(job_blob_dir):
    #             continue
    #         else:
    #             for ws_ in ws_pair:
    #                 pdf_print_formatting(ws_)
    #             wb.Worksheets(ws_pair).Select()
    #             wb.ActiveSheet.ExportAsFixedFormat(0, job_blob_dir + "\\" + f"{filename}.pdf")

    wb.Close(SaveChanges=False)
    excel_object.Quit()