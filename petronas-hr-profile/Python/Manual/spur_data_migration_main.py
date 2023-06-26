from numpy.core.numeric import False_
import pandas as pd
import numpy as np
import re
import win32com.client
import shutil
import os
import sys
import glob
import logging

import spur_data_processor
import spur_pd_processor
import spur_pptx_to_xlsx
import spur_xlsx_write_up_to_xlsx
import spur_job_profile
import spur_position_profile

# SKG Name
#######################################################################################################################
skg_name = "SKG009"
#######################################################################################################################

# Directory
main_dir = os.path.abspath(r"C:\Users\hradmin\Desktop\Data Migration\{}_SPUR_migration".format(skg_name))
data_dir = main_dir + "\\" + "data"
log_dir = main_dir + "\\" + "log_files"
job_dir = main_dir + "\\" + "Job_SPUR"
job_blob_dir = job_dir + "\\" + "BlobFiles"
job_clob_dir = job_dir + "\\" + "ClobFiles"
job_dat_dir = job_dir + "\\" + "DatFiles"
position_dir = main_dir + "\\" + "Position_SPUR"
position_blob_dir = position_dir + "\\" + "BlobFiles"
position_clob_dir = position_dir + "\\" + "ClobFiles"
position_dat_dir = position_dir + "\\" + "DatFiles"

if not os.path.exists(main_dir):
    os.makedirs(main_dir)
    os.makedirs(data_dir)
    # os.makedirs(data_dir + "\\" + "Competency")
    os.makedirs(data_dir + "\\" + "final_processed_data")
    os.makedirs(data_dir + "\\" + "Position_Master_Data")
    os.makedirs(data_dir + "\\" + "Simplified_Template")
    os.makedirs(data_dir + "\\" + "Extracted_data")
    os.makedirs(data_dir + "\\" + "Write_up")
    os.makedirs(data_dir + "\\" + "PD")
    os.makedirs(data_dir + "\\" + "JCP")
    os.makedirs(log_dir)
    os.makedirs(job_dir)
    os.makedirs(job_blob_dir)
    os.makedirs(job_clob_dir)
    os.makedirs(job_dat_dir)
    os.makedirs(position_dir)
    os.makedirs(position_blob_dir)
    os.makedirs(position_clob_dir)
    os.makedirs(position_dat_dir)

    shutil.copy2(
        r"data_migration\Oracle_template\PET_Job SPUR.xlsx",
        job_dir,
    )
    shutil.copy2(
        r"data_migration\Oracle_template\PET_Position SPUR.xlsx",
        position_dir,
    )

    sys.exit("{}_SPUR_migration folder created".format(skg_name))


# Code trigger
#######################################################################################################################
# Data processor
run_data_processor = False
run_pd_procesor = False

# PPTX write-up
run_pptx_extraction = True
save_slide = True

# XLSX write-up
run_xlsx_extraction = False
save_xlsx_sheet = False
save_xml = False

# Tasks
run_content_item_validation = False
run_job_profile = False
run_position_profile = False

# True if to read data from JCP, else False
read_jcp = False

# True if TC in simplified template, else False
tc_in_simplified_template = False
#######################################################################################################################

# Change files path accordingly
#######################################################################################################################
LC_file_path = r"data_migration\Competency_LC\Mapping LC for Position v1.xlsx"
content_item_file_path = r"data_migration\ContentItem\ContentItem Competency Prod 02102021.xlsx"
WS_path = r"data_migration\Work structure\ZHPLA 01112021 extract on 14092021 for Job Mapping to Business_latest compiled 07102021_Exclude E Series JG.xlsx"
#######################################################################################################################

# Ignore list
#######################################################################################################################
SPUR_ID_ignore_list = []  # ["F02-060", "F02-026", "15-001", "15-005"]
#######################################################################################################################


# Logging
log_dir_list = os.listdir(log_dir)
logging.basicConfig(
    filename=log_dir + "\\" + "code.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%d/%m/%Y %I:%M:%S %p",
    filemode="w",
)
stderrLogger = logging.StreamHandler()
stderrLogger.setFormatter(
    logging.Formatter(fmt="%(asctime)s | %(levelname)s | %(message)s", datefmt="%d/%m/%Y %I:%M:%S %p")
)
logging.getLogger().addHandler(stderrLogger)

# Processing data
if run_data_processor == True:
    # Simplified template data list
    simplified_template_file_path = data_dir + "\\" + "Simplified_Template\\*.xlsx"
    # Position master data
    position_master_data_file_path = glob.glob(data_dir + "\\" + "Position_Master_Data\\*.xlsx")[0]
    # position_master_data_file_path = ""
    # Company code mapping path
    cocode_map_path = glob.glob(r"data_migration\Company_code_map\*.xlsx")[0]

    if tc_in_simplified_template == True:
        tc_raw_path = simplified_template_file_path
    else:
        tc_raw_path = data_dir + "\\" + "Competency\\*.xlsx"

    # Processing simplified template data
    logging.info("[Processing data] Simplified template data")
    spur_data_processor.data_processor(
        main_dir=main_dir,
        position_master_data_file_path=position_master_data_file_path,
        tc_in_simplified_template=tc_in_simplified_template,
        simplified_template_file_path=simplified_template_file_path,
        tc_raw_path=tc_raw_path,
        cocode_map_path=cocode_map_path,
        skg_name=skg_name,
    ).simplified_template_data()
    # Position profile data
    logging.info("[Processing data] Position profile data")
    spur_data_processor.data_processor(
        main_dir=main_dir,
        position_master_data_file_path=position_master_data_file_path,
        tc_in_simplified_template=tc_in_simplified_template,
        simplified_template_file_path=simplified_template_file_path,
        tc_raw_path=tc_raw_path,
        cocode_map_path=cocode_map_path,
        skg_name=skg_name,
    ).position_profile_data()
    # # TC data
    logging.info("[Processing data] Technical competency data")
    spur_data_processor.data_processor(
        main_dir=main_dir,
        position_master_data_file_path=position_master_data_file_path,
        tc_in_simplified_template=tc_in_simplified_template,
        simplified_template_file_path=simplified_template_file_path,
        tc_raw_path=tc_raw_path,
        cocode_map_path=cocode_map_path,
        skg_name=skg_name,
    ).tc_data()

# Processing PD
if run_pd_procesor == True:
    # PD folder
    pd_folder = data_dir + "\\" + "PD" + "\\"
    spur_pd_processor.pd_processor(
        position_blob_dir=position_blob_dir,
        pd_folder=pd_folder,
    )

# Extract write-up
# if write_up_format == "pptx":
# Extract pptx write-up to xlsx
if run_pptx_extraction == True:
    ppt_list = glob.glob(data_dir + "\\" + "Write_up\**\**\*.pptx", recursive=True)
    ppt_list = list(set(ppt_list))
    pptx_df = spur_pptx_to_xlsx.pptx_to_xlsx(
        ppt_list=ppt_list,
        # xlsx_destination=main_dir + "\\" + "data\\Extracted_data\\{}_SPUR_Data_with_HTML.xlsx".format(skg_name),
        save_slide=save_slide,
        job_blob_dir=job_blob_dir,
        job_clob_dir=job_clob_dir,
        position_blob_dir=position_blob_dir,
        position_clob_dir=position_clob_dir,
    )
# else:
# Extract xlsx write-up to xlsx
if run_xlsx_extraction == True:
    xlsx_list = glob.glob(data_dir + "\\" + "Write_up\**\**\*.xlsx", recursive=True)
    xlsx_list = list(set(xlsx_list))
    xlsx_df = spur_xlsx_write_up_to_xlsx.xlsx_write_up_extract(
        xlsx_list=xlsx_list,
        # xlsx_destination=main_dir + "\\" + "data\\Extracted_data\\{}_SPUR_Data_with_HTML.xlsx".format(skg_name),
        save_xlsx_sheet=save_xlsx_sheet,
        save_xml=save_xml,
        data_dir=data_dir,
        job_blob_dir=job_blob_dir,
        job_clob_dir=job_clob_dir,
        position_blob_dir=position_blob_dir,
        position_clob_dir=position_clob_dir,
        skg_name=skg_name,
    )

if (run_pptx_extraction == True) & (run_xlsx_extraction == True):
    spur_df = pd.concat([pptx_df, xlsx_df]).reset_index(drop=True)
    spur_df = spur_df.sort_values(by=["UR_CODE"])
    spur_df.to_excel(
        main_dir + "\\" + "data\\Extracted_data\\{}_SPUR_Data_with_HTML.xlsx".format(skg_name), index=False
    )

elif (run_pptx_extraction == True) & (run_xlsx_extraction == False):
    pptx_df = pptx_df.sort_values(by=["UR_CODE"])
    pptx_df.to_excel(
        main_dir + "\\" + "data\\Extracted_data\\{}_SPUR_Data_with_HTML.xlsx".format(skg_name), index=False
    )

elif (run_pptx_extraction == False) & (run_xlsx_extraction == True):
    xlsx_df = xlsx_df.sort_values(by=["UR_CODE"])
    xlsx_df.to_excel(
        main_dir + "\\" + "data\\Extracted_data\\{}_SPUR_Data_with_HTML.xlsx".format(skg_name), index=False
    )

else:
    pass

# Files path
job_template_file_path = os.path.abspath(glob.glob(job_dir + "\\" + "*.xlsx")[0])
position_template_file_path = os.path.abspath(glob.glob(position_dir + "\\" + "*.xlsx")[0])
spur_data_file_path = main_dir + "\\" + "data\\Extracted_data\\{}_SPUR_Data_with_HTML.xlsx".format(skg_name)
spur_details_file_path = main_dir + "\\" + "data\\final_processed_data\\{}_details.xlsx".format(skg_name)
spur_position_file_path = (
    main_dir + "\\" + "data\\final_processed_data\\{}_position_profile_data.xlsx".format(skg_name)
)
TC_file_path = main_dir + "\\" + "data\\final_processed_data\\{}_TC.xlsx".format(skg_name)
# JCP data
if read_jcp == True:
    jcp_file_path = glob.glob(data_dir + "\\" + "JCP\\*.xlsx")[0]
else:
    jcp_file_path = ""


def content_item_validation(
    LC_file_path,
    TC_file_path,
    content_item_file_path,
    spur_data_file_path,
    spur_details_file_path,
    spur_position_file_path,
):
    content_item_dict = pd.read_excel(content_item_file_path, sheet_name=None)
    LC_content_item_list = (
        content_item_dict["ContentItem-Competency Edge"]
        .loc[9:, "Content Type Name"]
        .str.strip()
        .apply(lambda x: " ".join(str(x).split()))
    )
    # competency_tec_sheet = [x for x in content_item_dict.keys() if re.search("ContentItem-Competency$", x)][0]
    # TC_content_item_list = (
    #     content_item_dict[competency_tec_sheet]
    #     .loc[10:, "Content Type Name"]
    #     .str.strip()
    #     .apply(lambda x: " ".join(str(x).split()))
    # )
    membership_content_item_list = (
        content_item_dict["ContentItem-Membership"]
        .loc[10:, "Content Type Name"]
        .str.strip()
        .apply(lambda x: " ".join(str(x).split()))
    )
    awards_content_item_list = (
        content_item_dict["ContentItem-Honor & Awards"]
        .loc[10:, "Content Type Name"]
        .str.strip()
        .apply(lambda x: " ".join(str(x).split()))
    )
    license_content_item_list = (
        content_item_dict["ContentItem-License & Certif"]
        .loc[10:, "Content Type Name"]
        .str.strip()
        .apply(lambda x: " ".join(str(x).split()))
    )

    # Content item validation
    # LC validation
    competency_df = pd.read_excel(LC_file_path, sheet_name=None, header=1)
    lc_df = pd.concat([competency_df[list(competency_df.keys())[idx]] for idx in range(2, 7)]).reset_index(drop=True)
    # lc_df.iloc[:, 1:] = lc_df.iloc[:, 1:].astype("Int64")
    # remove double space
    lc_df["Sub-Competency"] = (
        lc_df["Sub-Competency"].apply(lambda x: " ".join(x.split())).str.replace("–", "-").str.replace(" : ", ": ")
    )
    LC_not_in_list = np.setdiff1d(lc_df["Sub-Competency"], [x for x in LC_content_item_list if isinstance(x, str)])
    if len(LC_not_in_list) != 0:
        logging.warning("[Content item validation] {} LC not in content item".format(len(LC_not_in_list)))
        with open(log_dir + "\\" + f"{skg_name}_LC_not_in_ContentItem.txt", "w") as f:
            f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(LC_not_in_list, start=1)))
    else:
        if f"{skg_name}_LC_not_in_ContentItem.txt" in log_dir_list:
            os.remove(log_dir + "\\" + f"{skg_name}_LC_not_in_ContentItem.txt")

    # TC validation
    # tc_df = pd.read_excel(TC_file_path)
    # oracle_column = [x for x in tc_df.columns if re.search("Oracle|Compentecy Technical|ContentItem", x, flags=re.I)][
    #     0
    # ]
    # tc_df[oracle_column] = (
    #     tc_df[oracle_column].apply(lambda x: " ".join(x.split()) if isinstance(x, str) else x).str.replace("–", "-")
    # )
    # if read_jcp == True:
    #     # jcp_df = pd.read_excel(jcp_file_path)
    #     jcp_df = pd.read_excel(jcp_file_path).replace("\u200b", "", regex=True)
    #     jcp_df["competency"] = jcp_df["Ti Name"].str.strip() + " " + jcp_df["Ti Number"].str.strip()
    #     jcp_df["competency"] = (
    #         jcp_df["competency"]
    #         .apply(lambda x: " ".join(x.split()) if isinstance(x, str) else x)
    #         .str.replace("–", "-")
    #     )
    #     tc_data_series = pd.concat([tc_df[oracle_column], jcp_df["competency"]])
    # else:
    #     tc_data_series = tc_df[oracle_column]
    # TC_not_in_list = np.setdiff1d(
    #     tc_data_series.astype(str),
    #     [x for x in TC_content_item_list if isinstance(x, str)],
    # )
    # if len(TC_not_in_list) != 0:
    #     logging.warning("[Content item validation] {} TC not in content item".format(len(TC_not_in_list)))
    #     with open(log_dir + "\\" + f"{skg_name}_TC_not_in_ContentItem.txt", "w") as f:
    #         f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(TC_not_in_list, start=1)))
    # else:
    #     if f"{skg_name}_TC_not_in_ContentItem.txt" in log_dir_list:
    #         os.remove(log_dir + "\\" + f"{skg_name}_TC_not_in_ContentItem.txt")

    # Awards validation
    awards_df = pd.read_excel(spur_details_file_path, sheet_name="Awards")
    awards_column = [x for x in awards_df.columns if re.search("Honor & Awards", x, flags=re.I)][0]
    awards_df[awards_column] = awards_df[awards_column].apply(lambda x: " ".join(x.split())).str.replace("–", "-")
    awards_not_in_list = np.setdiff1d(
        awards_df[awards_column],
        [x for x in awards_content_item_list if isinstance(x, str)],
    )
    if len(awards_not_in_list) != 0:
        logging.warning("[Content item validation] {} awards not in content item".format(len(awards_not_in_list)))
        with open(log_dir + "\\" + f"{skg_name}_awards_not_in_ContentItem.txt", "w") as f:
            f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(awards_not_in_list, start=1)))
    else:
        if f"{skg_name}_awards_not_in_ContentItem.txt" in log_dir_list:
            os.remove(log_dir + "\\" + f"{skg_name}_awards_not_in_ContentItem.txt")

    # membership validation
    membership_df = pd.read_excel(spur_details_file_path, sheet_name="Membership")
    membership_df_column = [x for x in membership_df.columns if re.search("Membership", x, flags=re.I)][0]
    membership_df[membership_df_column] = (
        membership_df[membership_df_column].apply(lambda x: " ".join(x.split())).str.replace("–", "-")
    )
    membership_not_in_list = np.setdiff1d(
        membership_df[membership_df_column],
        [x for x in membership_content_item_list if isinstance(x, str)],
    )
    if len(membership_not_in_list) != 0:
        logging.warning(
            "[Content item validation] {} membership not in content item".format(len(membership_not_in_list))
        )
        with open(log_dir + "\\" + f"{skg_name}_membership_not_in_ContentItem.txt", "w") as f:
            f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(membership_not_in_list, start=1)))
    else:
        if f"{skg_name}_membership_not_in_ContentItem.txt" in log_dir_list:
            os.remove(log_dir + "\\" + f"{skg_name}_membership_not_in_ContentItem.txt")

    # license validation
    license_df = pd.read_excel(spur_details_file_path, sheet_name="License").replace("\u200b", "", regex=True)
    license_df_column = [x for x in license_df.columns if re.search("License", x, flags=re.I)][0]
    license_df[license_df_column] = (
        license_df[license_df_column].apply(lambda x: " ".join(x.split())).str.replace("–", "-")
    )
    license_not_in_list = np.setdiff1d(
        license_df[license_df_column],
        [x for x in license_content_item_list if isinstance(x, str)],
    )
    if len(license_not_in_list) != 0:
        logging.warning("[Content item validation] {} License not in content item".format(len(license_not_in_list)))
        with open(log_dir + "\\" + f"{skg_name}_license_not_in_ContentItem.txt", "w") as f:
            f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(license_not_in_list, start=1)))
    else:
        if f"{skg_name}_license_not_in_ContentItem.txt" in log_dir_list:
            os.remove(log_dir + "\\" + f"{skg_name}_license_not_in_ContentItem.txt")

    # Data validation
    # TC validation
    spur_df = pd.read_excel(spur_data_file_path).sort_values("UR_CODE", ascending=True)
    spur_df = spur_df[~spur_df["UR_CODE"].isin(SPUR_ID_ignore_list)]
    TC_df = pd.read_excel(TC_file_path)
    spur_id_without_tc_list = np.setdiff1d(spur_df["UR_CODE"], TC_df["SPUR ID"])
    if len(spur_id_without_tc_list) != 0:
        logging.warning("[Data validation] {} SPUR ID have no TC data".format(len(spur_id_without_tc_list)))
        with open(log_dir + "\\" + f"{skg_name}_UR_without_TC.txt", "w") as f:
            f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(spur_id_without_tc_list, start=1)))
    else:
        if f"{skg_name}_UR_without_TC.txt" in log_dir_list:
            os.remove(log_dir + "\\" + f"{skg_name}_UR_without_TC.txt")

    # Position profile validation
    spur_df = pd.read_excel(spur_data_file_path).sort_values("UR_CODE", ascending=True)
    spur_df = spur_df[~spur_df["UR_CODE"].isin(SPUR_ID_ignore_list)]
    position_profile_df = pd.read_excel(spur_position_file_path)
    spur_id_without_pid_list = np.setdiff1d(spur_df["UR_CODE"], position_profile_df["SPUR ID"])
    if len(spur_id_without_pid_list) != 0:
        logging.warning("[Data validation] {} SPUR ID have no position data".format(len(spur_id_without_pid_list)))
        with open(log_dir + "\\" + f"{skg_name}_UR_without_PID.txt", "w") as f:
            f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(spur_id_without_pid_list, start=1)))
    else:
        if f"{skg_name}_UR_without_PID.txt" in log_dir_list:
            os.remove(log_dir + "\\" + f"{skg_name}_UR_without_PID.txt")

    # Write-up validation
    spur_df = pd.read_excel(spur_data_file_path).sort_values("UR_CODE", ascending=True)
    spur_df = spur_df[~spur_df["UR_CODE"].isin(SPUR_ID_ignore_list)]
    position_profile_df = pd.read_excel(spur_position_file_path)
    position_profile_df = position_profile_df[~position_profile_df["SPUR ID"].isin(SPUR_ID_ignore_list)]
    spur_id_without_write_up_list = np.setdiff1d(position_profile_df["SPUR ID"], spur_df["UR_CODE"])
    if len(spur_id_without_write_up_list) != 0:
        logging.warning("[Data validation] {} SPUR ID have no write up".format(len(spur_id_without_write_up_list)))
        with open(log_dir + "\\" + f"{skg_name}_UR_without_write_up.txt", "w") as f:
            f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(spur_id_without_write_up_list, start=1)))
    else:
        if f"{skg_name}_UR_without_write_up.txt" in log_dir_list:
            os.remove(log_dir + "\\" + f"{skg_name}_UR_without_write_up.txt")

    # Details validation
    experience_df = pd.read_excel(spur_details_file_path, sheet_name="Experience")
    degree_df = pd.read_excel(spur_details_file_path, sheet_name="Degree")
    membership_df = pd.read_excel(spur_details_file_path, sheet_name="Membership")
    awards_df = pd.read_excel(spur_details_file_path, sheet_name="Awards")

    ## Experience
    spur_without_exp_list = np.setdiff1d(spur_df["UR_CODE"], experience_df["SPUR ID"].dropna())
    if len(spur_without_exp_list) != 0:
        logging.warning("[Data validation] {} SPUR ID have no experience data".format(len(spur_without_exp_list)))
        with open(log_dir + "\\" + f"{skg_name}_SPUR_without_exp.txt", "w") as f:
            f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(spur_without_exp_list, start=1)))
    else:
        if f"{skg_name}_SPUR_without_exp.txt" in log_dir_list:
            os.remove(log_dir + "\\" + f"{skg_name}_SPUR_without_exp.txt")

    ## Degree
    degree_df = degree_df.dropna(subset=["SPUR ID"])
    spur_without_degree_list = np.setdiff1d(spur_df["UR_CODE"], degree_df["SPUR ID"])
    if len(spur_without_degree_list) != 0:
        logging.warning("[Data validation] {} SPUR ID have no degree data".format(len(spur_without_degree_list)))
        with open(log_dir + "\\" + f"{skg_name}_SPUR_without_degree.txt", "w") as f:
            f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(spur_without_degree_list, start=1)))
    else:
        if f"{skg_name}_SPUR_without_degree.txt" in log_dir_list:
            os.remove(log_dir + "\\" + f"{skg_name}_SPUR_without_degree.txt")


if run_content_item_validation == True:
    content_item_validation(
        LC_file_path,
        TC_file_path,
        content_item_file_path,
        spur_data_file_path,
        spur_details_file_path,
        spur_position_file_path,
    )

# Execute job profile
if run_job_profile == True:
    spur_job_profile.spur_job_profile(
        skg_name=skg_name,
        job_template_file_path=job_template_file_path,
        spur_data_file_path=spur_data_file_path,
        spur_details_file_path=spur_details_file_path,
        spur_position_file_path=spur_position_file_path,
        LC_file_path=LC_file_path,
        TC_file_path=TC_file_path,
        job_dat_dir=job_dat_dir,
        SPUR_ID_ignore_list=SPUR_ID_ignore_list,
        content_item_file_path=content_item_file_path,
        WS_path=WS_path,
        log_dir=log_dir,
    )

# Execute position profile
if run_position_profile == True:
    spur_position_profile.spur_position_profile(
        skg_name=skg_name,
        position_blob_dir=position_blob_dir,
        position_template_file_path=position_template_file_path,
        spur_data_file_path=spur_data_file_path,
        spur_details_file_path=spur_details_file_path,
        spur_position_file_path=spur_position_file_path,
        LC_file_path=LC_file_path,
        TC_file_path=TC_file_path,
        jcp_file_path=jcp_file_path,
        position_dat_dir=position_dat_dir,
        SPUR_ID_ignore_list=SPUR_ID_ignore_list,
        content_item_file_path=content_item_file_path,
        read_jcp=read_jcp,
        WS_path=WS_path,
        log_dir=log_dir,
    )

logging.info("Finished")