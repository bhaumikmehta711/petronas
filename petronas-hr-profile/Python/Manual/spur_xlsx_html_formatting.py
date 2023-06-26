import win32com.client
from xml.dom import minidom
from glob import glob
import itertools
import xml
import os
import re
import logging.config

log = logging.getLogger(__name__)


def save_excel_as_xml(xlsx_file, data_dir):
    excel_object = win32com.client.Dispatch("Excel.Application")
    excel_object.Visible = False
    excel_object.DisplayAlerts = False

    wb = excel_object.Workbooks.Open(xlsx_file)
    xml_filename = re.search(r"[^\\]+$", xlsx_file).group().replace(".xlsx", ".xml")
    log.info(f"[Save xlsx as xml] Generating {xml_filename}")
    xml_folder_path = data_dir + "\\Write_up\\XML_files"
    if not os.path.exists(xml_folder_path):
        os.makedirs(xml_folder_path)
    save_path = xml_folder_path + "\\" + xml_filename
    wb.SaveAs(
        Filename=save_path,
        FileFormat=46,
    )

    wb.Close()
    excel_object.Quit()


def get_bold_list(skg_name, xlsx_file):
    # xml_path = glob(
    #     r"D:\Project\HR_SPUR\data_migration\Oracle\{}_SPUR_migration\data\Write_up\XML_files\*xml".format(skg_name)
    # )

    # final_bold_list = []
    # for xml_file in xml_path:
    #     try:
    #         mydoc = minidom.parse(xml_file)
    #     except:
    #         continue
    filename = re.search(r"[^\\]+$", xlsx_file).group().replace(".xlsx", ".xml")
    xml_file = f"data_migration\Oracle\{skg_name}_SPUR_migration\data\Write_up\XML_files\{filename}"
    try:
        mydoc = minidom.parse(xml_file)
    except:
        return []
    bold_list = []
    for bold_text in mydoc.getElementsByTagName("B"):
        if type(bold_text.firstChild) == xml.dom.minidom.Element:
            try:
                bold_text = bold_text.getElementsByTagName("Font")[0].firstChild.nodeValue
            except:
                bold_text = bold_text.getElementsByTagName("U")[0].firstChild.nodeValue

        else:
            bold_text = bold_text.firstChild.nodeValue

        text = bold_text.replace("(", "\(").replace(")", "\)").replace("\xa0", "").replace(".", "\.").strip()
        bold_list.append(text)
    bold_list = list(set(bold_list))
    bold_list = [x for x in bold_list if x != "\\." and len(x) > 1]
    # final_bold_list.append(bold_list)

    bold_list = bold_list + [
        "Participate and share knowledge of best practices within the industry and higher learning institutions, e.g :",
        "Educational Qualification Degree",
        "Educational Qualification Field/Branch/Stream",
        "Years of Experience\(Overall\)",
        "Years of Experience\(Domain\)",
        "Domain",
        "Internal Certifications",
        "External License & Certifications",
        "Any Other Requirement",
        "General \(Malaysia & International Operations\)",
    ]
    bold_list = sorted(list(set(bold_list)), key=len, reverse=True)
    bold_list = ["\\b" + x + "\\b" for x in bold_list]
    return bold_list


def apply_html_format(text, bold_list):
    text = text.replace("\xa0", "").strip()
    for bold in bold_list:
        if re.search(bold, text):
            # print(bold)
            bold_text_location = re.search(bold, text).span(0)
            text = (
                text[: bold_text_location[0]]
                + "<strong>"
                + text[bold_text_location[0] : bold_text_location[1]]
                + "</strong>"
                + text[bold_text_location[1] :]
            )
        else:
            continue
    html_text = (
        text.replace("\n", "<br>")
        .replace("\t", " ")
        .replace("<br>Domain<br>", "<br><strong>Domain</strong><br>")
        .replace("<strong></strong>", "")
    )
    html_text = re.sub("(<br>){4,}", "<br>" * 3, html_text)
    return html_text