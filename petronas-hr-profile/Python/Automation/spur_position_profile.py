import pandas as pd
import numpy as np
import datetime
import re
import itertools
import xlwings as xw
import logging
import sys
import os


def spur_position_profile(
    skg_name,
    position_blob_dir,
    position_template_file_path,
    spur_data_file_path,
    spur_details_file_path,
    spur_position_file_path,
    LC_file_path,
    TC_file_path,
    jcp_file_path,
    position_dat_dir,
    SPUR_ID_ignore_list,
    content_item_file_path,
    read_jcp,
    WS_path,
    log_dir,
):
    run_all = False

    log = logging.getLogger(__name__)

    wb = xw.Book(position_template_file_path)
    spur_df = pd.read_excel(spur_data_file_path).sort_values("UR_CODE", ascending=True)
    spur_df = spur_df[~spur_df["UR_CODE"].isin(SPUR_ID_ignore_list)]
    position_profile_df = pd.read_excel(spur_position_file_path).sort_values("SPUR ID", ascending=True)
    # remove duplicates
    position_profile_df = position_profile_df.drop_duplicates()
    position_profile_df["Pos ID"] = position_profile_df["Pos ID"].astype(str).str.zfill(8)
    position_profile_obj = position_profile_df.select_dtypes(["object"])
    position_profile_df[position_profile_obj.columns] = position_profile_obj.apply(lambda x: x.str.strip())

    experience_df = pd.read_excel(spur_details_file_path, sheet_name="Experience").replace("\n", "<br>", regex=True)
    experience_df_obj = experience_df.select_dtypes(["object"])
    experience_df[experience_df_obj.columns] = experience_df_obj.apply(lambda x: x.str.strip())
    #experience_df = experience_df.drop(industry_column = None)
    industry_column = [x for x in experience_df.columns if re.search("Industry|Field", x, flags=re.I)][0]
    #industry_column = [x for x in experience_df.columns if re.search("Industry|Field", x).drop][0]
    #drop blank industry column
    #industry_column = industry_column.drop()
    try:
        domain_column = [x for x in experience_df.columns if re.search("Domain", x, flags=re.I)][0]
        domain_exist = True
    except:
        domain_exist = False
    min_years_column = [
        x
        for x in experience_df.columns
        if re.search(
            "Min Years|minimumExperienceRequired|mimimumExperienceRequired",
            x,
            flags=re.I,
        )
    ][0]
    max_years_column = [
        x for x in experience_df.columns if re.search("Max Years|Desired Years Of experience", x, flags=re.I)
    ][0]

    degree_df = pd.read_excel(spur_details_file_path, sheet_name="Degree")
    print(degree_df)
    degree_df_obj = degree_df.select_dtypes(["object"])
    degree_df[degree_df_obj.columns] = degree_df_obj.apply(lambda x: x.str.strip())
    degree_column = [x for x in degree_df.columns if re.search("Degree|ContentItem", str(x), flags=re.I)][0]
    area_of_study_column = [
        x for x in degree_df.columns if re.search("Area of study|AreaOfStudy", str(x), flags=re.I)
    ][0]
    # degree_df = degree_df.drop_duplicates(subset=["SPUR ID", degree_column, area_of_study_column], keep="first")

    def job_grade_map(job_grade):
        """
        docstring
        """
        if job_grade == "A1 - D1":
            return "A1&A2&A3&D1"
        elif job_grade == "A1 - D2":
            return "A1&A2&A3&D1&D2"
        elif job_grade == "D2 - M2":
            return "D2&D3&M1&M2"
        elif job_grade == "C1 - C2":
            return "C1&C2"
        else:
            return job_grade

    degree_df = degree_df.dropna(subset=["SPUR ID"])
    degree_df["JG"] = degree_df["JG"].apply(lambda x: job_grade_map(x))

    membership_df = pd.read_excel(spur_details_file_path, sheet_name="Membership")
    membership_df = membership_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    membership_column = [
        x
        for x in membership_df.columns
        if re.search(
            "Membership - Affiliation or Professional Body|Bodies membership",
            x,
            flags=re.I,
        )
    ][0]

    awards_df = pd.read_excel(spur_details_file_path, sheet_name="Awards")
    awards_df = awards_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    awards_column = [
        x
        for x in awards_df.columns
        if re.search(
            "Awards",
            x,
            flags=re.I,
        )
    ][0]

    license_df = pd.read_excel(spur_details_file_path, sheet_name="License")
    license_df_obj = license_df.select_dtypes(["object"])
    license_df[license_df_obj.columns] = license_df_obj.apply(lambda x: x.str.strip())
    license_column = [x for x in license_df.columns if re.search("License", x, flags=re.I)][0]

    content_item_dict = pd.read_excel(content_item_file_path, sheet_name=None)
    LC_content_item_list = (
        content_item_dict["ContentItem-Competency Edge"]
        .loc[9:, "Content Type Name"]
        .str.strip()
        .apply(lambda x: " ".join(str(x).split()))
    )
    # competency_tec_sheet = [
    #     x for x in content_item_dict.keys() if re.search("ContentItem-Competency$", x, flags=re.I)
    # ][0]
    # TC_content_item_list = (
    #     content_item_dict[competency_tec_sheet]
    #     .loc[10:, "Content Type Name"]
    #     .str.strip()
    #     .apply(lambda x: " ".join(str(x).split()))
    # )
    TC_content_item_list = []
    degree_sheet = [x for x in content_item_dict.keys() if re.search("^Degree|AreaOfStudy", x, flags=re.I)][0]
    AreaOfStudy_content_item_list = (
        content_item_dict[degree_sheet].iloc[1:, 2].str.strip().apply(lambda x: " ".join(x.split()))
    )
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

    WS_df = pd.read_excel(WS_path, sheet_name="Final for SPUR")
    WS_PID_list = WS_df.loc[:, "Pos ID (as per ZHPLA 01/11/2021)"].astype(str).apply(lambda x: x.split(".")[0]).str.zfill(8).tolist()
    log_dir_list = os.listdir(log_dir)

    def profile_code_map(string):
        """
        This function map role level of UR into designation
        """
        if "chief" in string.lower():
            return "SGM"

        elif "general manager" in string.lower() or "custodian" in string.lower() or "head" in string.lower():
            return "GM"

        elif "senior manager" in string.lower() or "principal" in string.lower() or "principle" in string.lower():
            return "SM"

        elif "manager" in string.lower() or "staff" in string.lower():
            return "MANAGER"

        elif "executive" in string.lower():
            return "EXECUTIVE"

        else:
            return ""

    ## Talent Profile ##
    def talent_profile(wb, position_profile_df, position_dat_dir):
        talent_profile = wb.sheets[0]
        talent_profile.range((12, "B"), (10000, "K")).clear()

        # Column D
        talent_profile.range((12, "D")).value = [[x] for x in position_profile_df["ProfileCode"].values.tolist()]
        row_end_talent_profile = talent_profile.range("D" + str(talent_profile.cells.last_cell.row)).end("up").row
        column_D = row_end_talent_profile

        # Column H
        talent_profile.range((12, "H")).value = [
            [x] for x in position_profile_df["Position"].replace("–", "-").values.tolist()
        ]

        # Column I
        talent_profile.range((12, "I"), (row_end_talent_profile, "I")).value = "=H12"

        # Column BCEFGJ
        for col in "BCEFGJ":
            talent_profile.range((12, col), (row_end_talent_profile, col)).value = "={}$11".format(col)

        talent_profile.range(
            (12, "K"), (row_end_talent_profile, "K")
        ).value = '=CONCAT(UPPER(SUBSTITUTE(D12," ","_")), "_POSITION_PROFILE")'

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(talent_profile.range("B1:K1").value).T
        talent_profile_df = pd.DataFrame(talent_profile.range("B12:K{}".format(row_end_talent_profile)).value)
        pid_missing_list = []
        profile_code_remove_list = []
        for profile_code in talent_profile_df[2]:
            pid = re.search("[^_]+$", profile_code).group()
            if pid not in WS_PID_list:
                pid_missing_list.append(pid)
                profile_code_remove_list.append(profile_code)
        if len(pid_missing_list) != 0:
            logging.warning("[Data validation] {} PID missing".format(len(pid_missing_list)))
            with open(log_dir + "\\" + f"{skg_name}_PID_missing.txt", "w") as f:
                f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(pid_missing_list, start=1)))
        else:
            if f"{skg_name}_PID_missing.txt" in log_dir_list:
                os.remove(log_dir + "\\" + f"{skg_name}_PID_missing.txt")
        talent_profile_df = talent_profile_df[~talent_profile_df[2].isin(profile_code_remove_list)]

        talent_profile.range((12, "B"), (10000, "K")).clear()
        talent_profile.range((12, "B")).value = talent_profile_df.values.tolist()
        talent_profile_df = pd.concat([header, talent_profile_df])
        # talent_profile_df = talent_profile_df.drop_duplicates(subset=["ProfileCode"])
        talent_profile_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_TalentProfile.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    ## Profile Relation ##
    def profile_relation(wb, position_dat_dir):
        profile_relation = wb.sheets[1]
        profile_relation.range((11, "B"), (10000, "K")).clear()
        row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row - 1

        # Column F
        profile_relation.range((11, "F"), (row_end_talent_profile, "F")).value = "=TalentProfile!D12"

        row_end_profile_relation = (
            profile_relation.range("F" + str(profile_relation.cells.last_cell.row)).end("up").row
        )

        # Column H
        profile_relation.range((11, "H")).value = [
            ["'" + re.search("[^_]+$", x).group().zfill(8)]
            for x in profile_relation.range((11, "F"), (row_end_profile_relation, "F")).value
        ]

        # Column I
        profile_relation.range((11, "I")).value = [
            [
                position_profile_df[position_profile_df["Pos ID"] == str(x).replace("'", "").zfill(8)][
                    "Company_full_name"
                ].iloc[0]
            ]
            for x in profile_relation.range((11, "H"), (row_end_profile_relation, "H")).value
        ]

        # Column BCDEGJ
        for col in "BCDEGJ":
            profile_relation.range((11, col), (row_end_profile_relation, col)).value = "={}$10".format(col)

        profile_relation.range(
            (11, "K"), (row_end_talent_profile, "K")
        ).value = '=CONCAT(UPPER(SUBSTITUTE(F11," ","_")), "_POSITION_PROFILE_RELATION")'

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(profile_relation.range("B1:K1").value).T
        profile_relation_df = pd.DataFrame(profile_relation.range("B11:K{}".format(row_end_talent_profile)).value)

        profile_relation_df[2] = "2021/11/01"
        profile_relation_df[3] = "4712/12/31"
        profile_relation_df[6] = profile_relation_df[6].astype(str).str.zfill(8)

        profile_relation.range((11, "B"), (10000, "K")).clear()
        profile_relation.range((11, "B")).value = profile_relation_df.values.tolist()

        profile_relation.range((11, "D"), (row_end_talent_profile, "D")).value = "'2021/11/01"
        profile_relation.range((11, "E"), (row_end_talent_profile, "E")).value = "'4712/12/31"
        # Column H
        profile_relation.range((11, "H")).value = [
            ["'" + re.search("[^_]+$", x).group().zfill(8)]
            for x in profile_relation.range((11, "F"), (row_end_profile_relation, "F")).value
        ]

        profile_relation_df = pd.concat([header, profile_relation_df])
        profile_relation_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ProfileRelation.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    def model_profile_info(wb, spur_df, position_dat_dir):
        ModelProfileExtraInfo = wb.sheets[2]
        ModelProfileExtraInfo.range((12, "B"), (100000, "I")).clear()
        row_end_talent_profile = wb.sheets[0].range("H" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

        # Column D
        ModelProfileExtraInfo.range((12, "D"), (row_end_talent_profile, "D")).value = "=TalentProfile!D12"
        row_end_ModelProfileExtraInfo = (
            ModelProfileExtraInfo.range("D" + str(ModelProfileExtraInfo.cells.last_cell.row)).end("up").row
        )

        # Column EFG
        ur_id_list = [
            re.sub("_.+", "", x) if re.sub("_.+", "", x) in spur_df["UR_CODE"].values.tolist() else ""
            for x in ModelProfileExtraInfo.range((12, "D")).expand("down").value
        ]
        ur_id_txt_list = [
            [
                x + "_DESCRIPTION.txt",
                x + "_QUALIFICATION.txt",
                x + "_RESPONSIBILITY.txt",
            ]
            if x in spur_df["UR_CODE"].values.tolist()
            else ["", "", ""]
            for x in ur_id_list
        ]
        ModelProfileExtraInfo.range((12, "E")).value = ur_id_txt_list

        # Column BCH
        for col in "BCH":
            ModelProfileExtraInfo.range((12, col), (row_end_ModelProfileExtraInfo, col)).value = "={}$11".format(col)

        ModelProfileExtraInfo.range(
            (12, "I"), (row_end_ModelProfileExtraInfo, "I")
        ).value = '=CONCAT(UPPER(SUBSTITUTE(D12," ","_")), "_JOB_PROFILE_MPEI")'

        for row in range(12, row_end_ModelProfileExtraInfo + 1):
            if ModelProfileExtraInfo.range((row, "E")).value == None:
                ModelProfileExtraInfo.range((row, "E")).color = (255, 0, 0)
                ModelProfileExtraInfo.range((row, "F")).color = (255, 0, 0)
                ModelProfileExtraInfo.range((row, "G")).color = (255, 0, 0)

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(ModelProfileExtraInfo.range("B1:I1").value).T
        ModelProfileExtraInfo_df = pd.DataFrame(
            ModelProfileExtraInfo.range("B12:I{}".format(row_end_ModelProfileExtraInfo)).value
        )
        ModelProfileExtraInfo_df = ModelProfileExtraInfo_df[
            (ModelProfileExtraInfo_df[3].notna()) | (ModelProfileExtraInfo_df[3] == "")
        ]

        ModelProfileExtraInfo.range((12, "B"), (10000, "I")).clear()
        ModelProfileExtraInfo.range((12, "B")).value = ModelProfileExtraInfo_df.values.tolist()

        ModelProfileExtraInfo_df = pd.concat([header, ModelProfileExtraInfo_df])
        # remove ModelProfileExtraInfo duplicates
        # ModelProfileExtraInfo_df = ModelProfileExtraInfo_df.drop_duplicates(subset=None, keep="first")
        ModelProfileExtraInfo_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ModelProfileExtraInfo.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    def profile_attachment(wb, spur_df, position_dat_dir, position_blob_dir):
        ProfileAttachment = wb.sheets[3]
        ProfileAttachment.range((11, "B"), (10000, "K")).clear()
        row_end_talent_profile = wb.sheets[0].range("H" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

        profile_code_list = []
        position_name_list = []
        final_blob_files_list = []

        blob_files_list = os.listdir(position_blob_dir)
        blob_files_name = [name.replace("PD", "").replace(".pdf", "").zfill(8) for name in blob_files_list]
        for row in range(12, row_end_talent_profile + 1):
            if wb.sheets[0].range((row, "D")).value == None:
                break

            profile_code = wb.sheets[0].range((row, "D")).value
            position_name = wb.sheets[0].range((row, "H")).get_address(False, False, True)
            ur_id = re.search("[^_]+", profile_code).group()
            pid = re.search("[^_]+$", profile_code).group().zfill(8)

            if pid in blob_files_name:
                profile_code_list.extend(["=" + wb.sheets[0].range((row, "D")).get_address(False, False, True)] * 2)
                position_name_list.extend(["=" + position_name] * 2)
                final_blob_files_list.extend(["{}.pdf".format(ur_id), "{}.pdf".format(pid)])

            else:
                profile_code_list.extend(["=" + wb.sheets[0].range((row, "D")).get_address(False, False, True)])
                position_name_list.extend(["=" + position_name])
                final_blob_files_list.extend(["{}.pdf".format(ur_id)])

        # Column H
        ProfileAttachment.range((11, "H")).value = [[x] for x in profile_code_list]

        row_end_ProfileAttachment = (
            ProfileAttachment.range("H" + str(ProfileAttachment.cells.last_cell.row)).end("up").row
        )

        # Column D
        ProfileAttachment.range((11, "D")).value = [[x] for x in position_name_list]

        # Column E
        ProfileAttachment.range((11, "E")).value = [[x] for x in final_blob_files_list]

        # Column F
        ProfileAttachment.range((11, "F"), (row_end_ProfileAttachment, "F")).value = "=D11"

        # Column I
        ProfileAttachment.range((11, "I"), (row_end_ProfileAttachment, "I")).value = '=E11 & ""'

        # # Column H
        # ProfileAttachment.range(
        #     (12, "H"), (row_end_talent_profile, "H")
        # ).value = "=TalentProfile!D12"

        # # Column D
        # ProfileAttachment.range(
        #     (12, "D"), (row_end_ProfileAttachment, "D")
        # ).value = "=TalentProfile!H12"

        # # Column F
        # ProfileAttachment.range(
        #     (12, "F"), (row_end_ProfileAttachment, "F")
        # ).value = "=D12"

        # # Column E
        # ur_id_list = [
        #     re.sub("_.+", "", x)
        #     if re.sub("_.+", "", x) in spur_df["UR_CODE"].values.tolist()
        #     else ""
        #     for x in ProfileAttachment.range((12, "H")).expand("down").value
        # ]
        # ur_id_pptx_list = [
        #     [x + ".pdf"] if x in spur_df["UR_CODE"].values.tolist() else [""]
        #     for x in ur_id_list
        # ]
        # ProfileAttachment.range((12, "E")).value = ur_id_pptx_list

        # # Column I
        # ProfileAttachment.range(
        #     (12, "I"), (row_end_ProfileAttachment, "I")
        # ).value = '=E12 & ""'

        # Column BCGJ
        # for col in "BCGJ":
        #     ProfileAttachment.range((11, col), (row_end_ProfileAttachment, col)).value = "={}$10".format(col)
        ProfileAttachment.range((11, "B"), (row_end_ProfileAttachment, "B")).value = "MERGE"
        ProfileAttachment.range((11, "C"), (row_end_ProfileAttachment, "C")).value = "ProfileAttachment"
        ProfileAttachment.range((11, "G"), (row_end_ProfileAttachment, "G")).value = "FILE"
        ProfileAttachment.range((11, "J"), (row_end_ProfileAttachment, "J")).value = "PETRONAS"

        # Column K
        z_list = [
            (
                ProfileAttachment.range((i, "H")).formula.replace("=", ""),
                ProfileAttachment.range((i, "E")).value,
            )
            for i in range(11, row_end_ProfileAttachment + 1)
        ]
        formula_list = [['=CONCATENATE({},"_",UPPER("{}"),"_ATTACHMENT ")'.format(k[0], k[1])] for k in z_list]

        ProfileAttachment.range((11, "K"), (row_end_ProfileAttachment, "K")).value = formula_list

        for row in range(11, row_end_ProfileAttachment + 1):
            if ProfileAttachment.range((row, "E")).value == None:
                ProfileAttachment.range((row, "E")).color = (255, 0, 0)
                ProfileAttachment.range((row, "I")).color = (255, 0, 0)

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(ProfileAttachment.range("B1:K1").value).T
        ProfileAttachment_df = pd.DataFrame(ProfileAttachment.range("B11:K{}".format(row_end_ProfileAttachment)).value)
        ProfileAttachment_df.loc[ProfileAttachment_df[3].str.contains("\d{8}"), 2] += " PD"
        # ProfileAttachment_df = ProfileAttachment_df[
        #     (ProfileAttachment_df[3].notna()) | (ProfileAttachment_df[3] == "")
        # ]
        # ProfileAttachment.range((12, "B"), (10000, "K")).clear()
        # ProfileAttachment.range((12, "B")).value = ProfileAttachment_df.values.tolist()

        ProfileAttachment.range((11, "B"), (10000, "K")).clear()
        ProfileAttachment.range((11, "B")).value = ProfileAttachment_df.values.tolist()

        ProfileAttachment_df = pd.concat([header, ProfileAttachment_df])
        ProfileAttachment_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ProfileAttachment.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    def profile_item_other_descriptor(wb, spur_df, position_dat_dir):
        ProfileItem_OtherDescriptor = wb.sheets[7]
        ProfileItem_OtherDescriptor.range((9, "B"), (10000, "Q")).clear()
        row_end_talent_profile = wb.sheets[1].range("H" + str(wb.sheets[1].cells.last_cell.row)).end("up").row

        # Column H
        ProfileItem_OtherDescriptor.range((9, "H")).value = [
            ["={}".format(wb.sheets[1].range((row, "K")).get_address(False, False, True))]
            for row in range(12, row_end_talent_profile + 1)
        ]
        row_end_ProfileItem_OtherDescriptor = (
            ProfileItem_OtherDescriptor.range("H" + str(ProfileItem_OtherDescriptor.cells.last_cell.row)).end("up").row
        )

        # Column BCDFGJKP
        for col in "BCDJKP":
            ProfileItem_OtherDescriptor.range(
                (9, col), (row_end_ProfileItem_OtherDescriptor, col)
            ).value = "={}$8".format(col)

        # Column F
        ProfileItem_OtherDescriptor.range((9, "F"), (row_end_ProfileItem_OtherDescriptor, "F")).value = "'2021/11/01"

        # Column G
        ProfileItem_OtherDescriptor.range((9, "G"), (row_end_ProfileItem_OtherDescriptor, "G")).value = "'4712/12/31 &"

        # Column L
        ProfileItem_OtherDescriptor.range((9, "L"), (row_end_ProfileItem_OtherDescriptor, "L")).value = [
            [
                spur_df[spur_df["UR_CODE"] == re.sub("_.+", "", x)]["CHALLENGES"].values[0].replace("\n", "")
                if re.sub("_.+", "", x) in spur_df["UR_CODE"].values.tolist()
                else np.nan
            ]
            for x in ProfileItem_OtherDescriptor.range("H9").expand("down").value
        ]

        # Column I
        ProfileItem_OtherDescriptor.range((9, "I"), (row_end_ProfileItem_OtherDescriptor, "I")).value = 5

        # Column Q
        ProfileItem_OtherDescriptor.range(
            (9, "Q"), (row_end_ProfileItem_OtherDescriptor, "Q")
        ).value = '=CONCAT(UPPER(SUBSTITUTE(TalentProfile!D12," ","_")),"_",I9,"_",UPPER(K9),"_PI")'

        for row in range(9, row_end_ProfileItem_OtherDescriptor + 1):
            if ProfileItem_OtherDescriptor.range((row, "L")).value == None:
                ProfileItem_OtherDescriptor.range((row, "L")).color = (255, 0, 0)

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(ProfileItem_OtherDescriptor.range("B2:Q2").value).T
        ProfileItem_OtherDescriptor_df = pd.DataFrame(
            ProfileItem_OtherDescriptor.range("B9:Q{}".format(row_end_ProfileItem_OtherDescriptor)).value
        )
        # print(ProfileItem_OtherDescriptor_df[10].notna())
        ProfileItem_OtherDescriptor_df = ProfileItem_OtherDescriptor_df[
            (ProfileItem_OtherDescriptor_df[10].notna()) | (ProfileItem_OtherDescriptor_df[10] == "")
        ]
        ProfileItem_OtherDescriptor_df[4] = "2021/11/01"
        ProfileItem_OtherDescriptor_df[5] = "4712/12/31"
        ProfileItem_OtherDescriptor_df[[7, 8]] = ProfileItem_OtherDescriptor_df[[7, 8]].apply(
            pd.to_numeric, downcast="signed"
        )

        ProfileItem_OtherDescriptor.range((9, "B"), (10000, "Q")).clear()
        ProfileItem_OtherDescriptor.range((9, "B")).value = ProfileItem_OtherDescriptor_df.values.tolist()

        ProfileItem_OtherDescriptor.range((9, "F"), (row_end_ProfileItem_OtherDescriptor, "F")).value = "'2021/11/01"
        ProfileItem_OtherDescriptor.range((9, "G"), (row_end_ProfileItem_OtherDescriptor, "G")).value = "'4712/12/31"

        ProfileItem_OtherDescriptor_df = pd.concat([header, ProfileItem_OtherDescriptor_df])
        ProfileItem_OtherDescriptor_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ProfileItem-OtherDescriptor.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    def profile_item_risk(wb, spur_df, position_dat_dir):
        ProfileItem_Risk = wb.sheets[25]
        ProfileItem_Risk.range((12, "B"), (10000, "L")).clear()
        row_end_talent_profile = wb.sheets[0].range("H" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

        # Column G
        ProfileItem_Risk.range((12, "G")).value = [
            ["={}".format(wb.sheets[0].range((row, "K")).get_address(False, False, True))]
            for row in range(12, row_end_talent_profile + 1)
        ]
        row_end_ProfileItem_Risk = (
            ProfileItem_Risk.range("G" + str(ProfileItem_Risk.cells.last_cell.row)).end("up").row
        )

        # Column BCDEFHIK
        for col in "BCDEFHIK":
            ProfileItem_Risk.range((12, col), (row_end_ProfileItem_Risk, col)).value = "={}$11".format(col)

        # Column E
        ProfileItem_Risk.range((12, "E"), (row_end_ProfileItem_Risk, "F")).value = "'2021/11/01"

        # Column F
        ProfileItem_Risk.range((12, "F"), (row_end_ProfileItem_Risk, "F")).value = "'4712/12/31 &"

        # Column J
        ProfileItem_Risk.range((12, "J"), (row_end_ProfileItem_Risk, "J")).value = [
            [
                spur_df[spur_df["UR_CODE"] == re.sub("_.+", "", x)]["CHALLENGES"].values[0].replace("\n", "")
                if re.sub("_.+", "", x) in spur_df["UR_CODE"].values.tolist()
                else np.nan
            ]
            for x in ProfileItem_Risk.range("G12").expand("down").value
        ]

        # Column L
        ProfileItem_Risk.range(
            (12, "L"), (row_end_ProfileItem_Risk, "L")
        ).value = '=CONCATENATE(UPPER(TalentProfile!D12),"_",UPPER(SUBSTITUTE(H12," ","_")),"_PI")'

        for row in range(12, row_end_ProfileItem_Risk + 1):
            if ProfileItem_Risk.range((row, "L")).value == None:
                ProfileItem_Risk.range((row, "L")).color = (255, 0, 0)

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(ProfileItem_Risk.range("B2:L2").value).T
        ProfileItem_Risk_df = pd.DataFrame(ProfileItem_Risk.range("B12:L{}".format(row_end_ProfileItem_Risk)).value)
        # print(ProfileItem_OtherDescriptor_df[10].notna())
        # ProfileItem_Risk_df = ProfileItem_Risk_df[(ProfileItem_Risk_df[10].notna()) | (ProfileItem_Risk_df[10] == "")]
        ProfileItem_Risk_df[3] = "2021/11/01"
        ProfileItem_Risk_df[4] = "4712/12/31"

        ProfileItem_Risk.range((12, "B"), (10000, "L")).clear()
        ProfileItem_Risk.range((12, "B")).value = ProfileItem_Risk_df.values.tolist()

        ProfileItem_Risk.range((12, "E"), (row_end_ProfileItem_Risk, "E")).value = "'2021/11/01"
        ProfileItem_Risk.range((12, "F"), (row_end_ProfileItem_Risk, "F")).value = "'4712/12/31"

        ProfileItem_OtherDescriptor_df = pd.concat([header, ProfileItem_Risk_df])
        ProfileItem_OtherDescriptor_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ProfileItem-Risk.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    def profile_item_exp_required(wb, experience_df, position_profile_df, position_dat_dir):
        ProfileItem_ExperienceRequired = wb.sheets[16]
        ProfileItem_ExperienceRequired.range((12, "B"), (100000, "Q")).clear()
        spur_id_occ = experience_df["SPUR ID"].value_counts().sort_index()
        #spur_id_occ = spur_id_occ[spur_id_occ != 1]

        id_list = []
        exp_importance_list = []
        min_exp_list = []
        max_exp_list = []
        industry_list = []
        domain_list = []

        row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
        for row in range(12, row_end_talent_profile + 1):
            if wb.sheets[0].range((row, "K")).value == None:
                break

            ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
            ur_id_exp_df = experience_df[experience_df["SPUR ID"] == ur_id]
            position_id = str(re.search("[^_]+$", wb.sheets[0].range((row, "D")).value).group()).zfill(8)
            role_level_column = [
                x for x in position_profile_df.columns if re.search("role level|role", x.lower().strip())
            ][0]
            role_level = position_profile_df[position_profile_df["Pos ID"] == position_id][role_level_column].values[0]
            jg_column = [x for x in position_profile_df.columns if "jg" in x.lower()][0]
            conso_JG = position_profile_df[position_profile_df["Pos ID"] == position_id][jg_column].values[0]
            # conso_JG = WS_df[WS_df["Pos ID (as per ZHPLA 3/5/2021)"] == position_id]["Conso JG"].iloc[0]
            conso_JG = re.sub(r"Est.|Eqv.", r"", conso_JG).strip()

            if ur_id in spur_id_occ.index:
                if all(ur_id_exp_df["JG"].notna()) and any([str(conso_JG) in x for x in ur_id_exp_df["JG"]]):
                    ur_id_exp_df = ur_id_exp_df.iloc[[conso_JG in x for x in ur_id_exp_df["JG"]]]
                    data = []
                    spur_id = ur_id
                    for field in ur_id_exp_df[industry_column].unique():
                        if field is np.nan:
                            min_years = ur_id_exp_df[ur_id_exp_df[industry_column].isna()][min_years_column].min()
                            max_years = ur_id_exp_df[ur_id_exp_df[industry_column].isna()][max_years_column].max()
                        else:
                            min_years = ur_id_exp_df[ur_id_exp_df[industry_column] == field][min_years_column].min()
                            max_years = ur_id_exp_df[ur_id_exp_df[industry_column] == field][max_years_column].max()
                        if domain_exist == True:
                            try:
                                domain = ur_id_exp_df[ur_id_exp_df[industry_column] == field][domain_column].iloc[0]
                            except:
                                domain = ur_id_exp_df[ur_id_exp_df[industry_column].isna()][domain_column].iloc[0]
                        try:
                            importance = ur_id_exp_df[ur_id_exp_df[industry_column] == field]["Importance"].iloc[0]
                        except:
                            importance = ur_id_exp_df[ur_id_exp_df[industry_column].isna()]["Importance"].iloc[0]
                        if domain_exist == True:
                            data.append(
                                [
                                    spur_id,
                                    min_years,
                                    max_years,
                                    field,
                                    domain,
                                    conso_JG,
                                    importance,
                                ]
                            )
                        else:
                            data.append(
                                [
                                    spur_id,
                                    min_years,
                                    max_years,
                                    field,
                                    conso_JG,
                                    importance,
                                ]
                            )
                    if domain_exist == True:
                        ur_id_exp_df = pd.DataFrame(
                            data,
                            columns=[
                                "SPUR ID",
                                min_years_column,
                                max_years_column,
                                industry_column,
                                domain_column,
                                "JG",
                                "Importance",
                            ],
                        )
                    else:
                        ur_id_exp_df = pd.DataFrame(
                            data,
                            columns=[
                                "SPUR ID",
                                min_years_column,
                                max_years_column,
                                industry_column,
                                "JG",
                                "Importance",
                            ],
                        )
                    id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(False, False, True)] * len(ur_id_exp_df)
                    )
                else:
                    data = []
                    spur_id = ur_id
                    for field in ur_id_exp_df[industry_column].unique():
                        if field is np.nan:
                            min_years = ur_id_exp_df[ur_id_exp_df[industry_column].isna()][min_years_column].min()
                            max_years = ur_id_exp_df[ur_id_exp_df[industry_column].isna()][max_years_column].max()
                        else:
                            min_years = ur_id_exp_df[ur_id_exp_df[industry_column] == field][min_years_column].min()
                            max_years = ur_id_exp_df[ur_id_exp_df[industry_column] == field][max_years_column].max()
                        if domain_exist == True:
                            try:
                                domain = ur_id_exp_df[ur_id_exp_df[industry_column] == field][domain_column].iloc[0]
                            except:
                                domain = ur_id_exp_df[ur_id_exp_df[industry_column].isna()][domain_column].iloc[0]
                        try:
                            importance = ur_id_exp_df[ur_id_exp_df[industry_column] == field]["Importance"].iloc[0]
                        except:
                            importance = ur_id_exp_df[ur_id_exp_df[industry_column].isna()]["Importance"].iloc[0]
                        if domain_exist == True:
                            data.append(
                                [
                                    spur_id,
                                    min_years,
                                    max_years,
                                    field,
                                    domain,
                                    conso_JG,
                                    importance,
                                ]
                            )
                        else:
                            data.append(
                                [
                                    spur_id,
                                    min_years,
                                    max_years,
                                    field,
                                    conso_JG,
                                    importance,
                                ]
                            )
                    if domain_exist == True:
                        ur_id_exp_df = pd.DataFrame(
                            data,
                            columns=[
                                "SPUR ID",
                                min_years_column,
                                max_years_column,
                                industry_column,
                                domain_column,
                                "JG",
                                "Importance",
                            ],
                        )
                    else:
                        ur_id_exp_df = pd.DataFrame(
                            data,
                            columns=[
                                "SPUR ID",
                                min_years_column,
                                max_years_column,
                                industry_column,
                                "JG",
                                "Importance",
                            ],
                        )
                    id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(False, False, True)] * len(ur_id_exp_df)
                    )

            else:
                if run_all == True:
                    id_list.extend([wb.sheets[0].range((row, "K")).get_address(False, False, True)])
                else:
                    continue

            if ur_id in experience_df["SPUR ID"].values.tolist():
                if all(ur_id_exp_df["JG"].notna()) and any([conso_JG in x for x in ur_id_exp_df["JG"]]):
                    exp_importance_list.extend(
                        ur_id_exp_df[ur_id_exp_df["SPUR ID"] == ur_id][["Importance"]].values.tolist()
                    )
                    min_exp_list.extend(ur_id_exp_df[[min_years_column]].values.tolist())
                    max_exp_list.extend(ur_id_exp_df[[max_years_column]].values.tolist())
                    industry_list.extend(ur_id_exp_df[[industry_column]].values.tolist())
                    if domain_exist == True:
                        domain_list.extend(ur_id_exp_df[[domain_column]].values.tolist())
                    else:
                        domain_list.extend(ur_id_exp_df[[industry_column]].values.tolist())
                else:
                    exp_importance_list.extend(
                        ur_id_exp_df[ur_id_exp_df["SPUR ID"] == ur_id][["Importance"]].values.tolist()
                    )
                    min_exp_list.extend(
                        ur_id_exp_df[ur_id_exp_df["SPUR ID"] == ur_id][[min_years_column]].values.tolist()
                    )
                    max_exp_list.extend(
                        ur_id_exp_df[ur_id_exp_df["SPUR ID"] == ur_id][[max_years_column]].values.tolist()
                    )
                    industry_list.extend(
                        ur_id_exp_df[ur_id_exp_df["SPUR ID"] == ur_id][[industry_column]].values.tolist()
                    )
                    if domain_exist == True:
                        domain_list.extend(
                            ur_id_exp_df[ur_id_exp_df["SPUR ID"] == ur_id][[domain_column]].values.tolist()
                        )
                    else:
                        domain_list.extend(
                            ur_id_exp_df[ur_id_exp_df["SPUR ID"] == ur_id][[industry_column]].values.tolist()
                        )

            else:
                if run_all == True:
                    exp_importance_list.extend([[""]])
                    min_exp_list.extend([[""]])
                    max_exp_list.extend([[""]])
                    industry_list.extend([[""]])
                    domain_list.extend([[""]])
                else:
                    continue

        # Column H
        ProfileItem_ExperienceRequired.range((12, "H")).value = [["={}".format(k)] for k in id_list]
        row_end_ProfileItem_ExperienceRequired = (
            ProfileItem_ExperienceRequired.range("H" + str(ProfileItem_ExperienceRequired.cells.last_cell.row))
            .end("up")
            .row
        )

        # Column J
        #ProfileItem_ExperienceRequired.range((12, "J")).value = jg_column

        # Column L
        ProfileItem_ExperienceRequired.range((12, "L")).value = min_exp_list

        # Column M
        ProfileItem_ExperienceRequired.range((12, "M")).value = max_exp_list

        # Column N
        ProfileItem_ExperienceRequired.range((12, "N")).value = domain_list

        # Column O
        ProfileItem_ExperienceRequired.range((12, "O")).value = industry_list

        # Column F
        ProfileItem_ExperienceRequired.range(
            (12, "F"), (row_end_ProfileItem_ExperienceRequired, "F")
        ).value = "'2021/11/01"

        # Column G
        ProfileItem_ExperienceRequired.range(
            (12, "G"), (row_end_ProfileItem_ExperienceRequired, "G")
        ).value = "'4712/12/31"

        # Column BCDFGIJKP
        for k in "BCDFGIJKP":
            ProfileItem_ExperienceRequired.range(
                "{}12:{}{}".format(k, k, row_end_ProfileItem_ExperienceRequired)
            ).value = "={}$11".format(k)

        # Column Q
        z_list = [
            (
                ProfileItem_ExperienceRequired.range((i, "H"))
                .formula.replace("=", "")
                .replace("TalentProfile!K", "TalentProfile!D"),
                ProfileItem_ExperienceRequired.range((i, "I")).get_address(False, False),
                ProfileItem_ExperienceRequired.range((i, "O")).get_address(False, False),
                ProfileItem_ExperienceRequired.range((i, "J")).get_address(False, False),
            )
            for i in range(12, row_end_ProfileItem_ExperienceRequired + 1)
        ]
        formula_list = [
            [
                '=CONCAT(UPPER(SUBSTITUTE({}," ","_")),"_",{},"_",UPPER(SUBSTITUTE({}," ","_")),"_",UPPER(SUBSTITUTE({}, " " , "_")) ,"_PI")'.format(
                    k[0], k[1], k[2], k[3]
                )
            ]
            for k in z_list
        ]
        ProfileItem_ExperienceRequired.range((12, "Q")).value = formula_list

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(ProfileItem_ExperienceRequired.range("B2:Q2").value).T
        ProfileItem_ExperienceRequired_df = pd.DataFrame(
            ProfileItem_ExperienceRequired.range("B12:Q{}".format(row_end_ProfileItem_ExperienceRequired)).value
        )
        ProfileItem_ExperienceRequired_df[4] = "2021/11/01"
        ProfileItem_ExperienceRequired_df[5] = "4712/12/31"
        ProfileItem_ExperienceRequired_df[[10, 11]] = ProfileItem_ExperienceRequired_df[[10, 11]].apply(
            pd.to_numeric, downcast="signed"
        )

        ProfileItem_ExperienceRequired_df[10] = (
            ProfileItem_ExperienceRequired_df[10].fillna(-1).astype(int).replace(-1, "")
        )
        ProfileItem_ExperienceRequired_df[11] = (
            ProfileItem_ExperienceRequired_df[11].fillna(-1).astype(int).replace(-1, "")
        )

        ProfileItem_ExperienceRequired.range((12, "B"), (10000, "Q")).clear()
        ProfileItem_ExperienceRequired.range((12, "B")).value = ProfileItem_ExperienceRequired_df.values.tolist()
        row_end_ProfileItem_ExperienceRequired_2 = (
            ProfileItem_ExperienceRequired.range("H" + str(ProfileItem_ExperienceRequired.cells.last_cell.row))
            .end("up")
            .row
        )
        ProfileItem_ExperienceRequired.range(
            (12, "F"), (row_end_ProfileItem_ExperienceRequired_2, "F")
        ).value = "'2021/11/01"
        ProfileItem_ExperienceRequired.range(
            (12, "G"), (row_end_ProfileItem_ExperienceRequired_2, "G")
        ).value = "'4712/12/31"

        ProfileItem_ExperienceRequired_df = pd.concat([header, ProfileItem_ExperienceRequired_df])
        # ProfileItem_ExperienceRequired_df = ProfileItem_ExperienceRequired_df.drop_duplicates(subset=[16], keep='first')
        ProfileItem_ExperienceRequired_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ProfileItem-ExperienceRequired.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

        for row in range(12, row_end_ProfileItem_ExperienceRequired + 1):
            if ProfileItem_ExperienceRequired.range((row, "O")).value == None:
                # ProfileItem_ExperienceRequired.range((row, 'J')).color = (255, 0, 0)
                # ProfileItem_ExperienceRequired.range((row, "M")).color = (255, 0, 0)
                # ProfileItem_ExperienceRequired.range((row, "N")).color = (255, 0, 0)
                ProfileItem_ExperienceRequired.range((row, "O")).color = (255, 0, 0)
                # ProfileItem_ExperienceRequired.range((row, "P")).color = (255, 0, 0)
            if ProfileItem_ExperienceRequired.range((row, "M")).value == None:
                ProfileItem_ExperienceRequired.range((row, "M")).color = (255, 0, 0)
            if ProfileItem_ExperienceRequired.range((row, "N")).value == None:
                ProfileItem_ExperienceRequired.range((row, "N")).color = (255, 0, 0)
            if ProfileItem_ExperienceRequired.range((row, "P")).value == None:
                ProfileItem_ExperienceRequired.range((row, "P")).color = (255, 0, 0)

    def profile_item_competency_LC(
        wb, LC_file_path, spur_df, position_profile_df, position_dat_dir, LC_content_item_list
    ):
        ProfileItem_Competency_LC = wb.sheets[20]
        ProfileItem_Competency_LC.range((13, "B"), (100000, "O")).clear()

        role_level_map = dict(zip(spur_df["UR_CODE"], spur_df["UR_NAME"].apply(profile_code_map)))
        sheets = [
            "Executive",
            "Manager",
            "Senior Managers",
            "Staff",
            "Principal",
            "Custodian",
            "General Managers",
            "Senior General Manager ++",
        ]
        competency_lc_dict = {}
        min_lc_dict = {}
        max_lc_dict = {}

        # lc_dfs = []
        # for sheet in sheets:
        #     lc_df = pd.read_excel(LC_file_path, sheet_name=sheet, header=1).dropna(axis=1)

        #     # remove double space
        #     lc_df["Sub-Competency"] = (
        #         lc_df["Sub-Competency"].apply(lambda x: " ".join(x.split())).str.replace(" : ", ": ")
        #     )
        #     lc_dfs.append(lc_df)

        # final_lc_df = pd.concat(lc_dfs, axis=1)
        # final_lc_df = final_lc_df.loc[:, ~final_lc_df.columns.duplicated()]

        lc_list = []
        min_list = []
        max_list = []

        row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
        for ur in [
            wb.sheets[0].range((row, 4)).value
            for row in range(12, row_end_talent_profile + 1)
            if wb.sheets[0].range((row, 4)).value != None
        ]:
            # ur = re.sub("_.+", "", ur)
            pid = re.search("[^_]+$", ur).group().zfill(8)
            role_level_column = [
                x for x in position_profile_df.columns if re.search("role level|role$", x.lower().strip())
            ][0]
            role_level = position_profile_df[position_profile_df["Pos ID"] == str(pid)][role_level_column].values[0]
            job_grade = position_profile_df[position_profile_df["Pos ID"] == str(pid)]["JG"].iloc[0]
            # if not isinstance(job_grade, str):
            # job_grade = WS_df[WS_df["Pos ID (as per ZHPLA 3/5/2021)"] == int(pid)]["Conso JG"].iloc[0]
            job_grade = job_grade.replace("Eqv.", "").replace("Est.", "").strip()

            if role_level in ["Staff", "Principal", "Custodian"]:
                sheet_name = role_level
                if role_level == "Staff":
                    job_grade = "E3"
                elif role_level == "Principal":
                    job_grade = "E4"
                else:
                    job_grade = "E5"
            else:
                if re.search("H1|H2", job_grade):
                    sheet_name = "Senior General Manager ++"

                if re.search("C1|C2", job_grade):
                    sheet_name = "General Managers"

                if re.search("M1|M2", job_grade):
                    sheet_name = "Senior Managers"

                if re.search("D2|D3", job_grade):
                    sheet_name = "Manager"

                if re.search("A1|A2|A3|D1", job_grade):
                    sheet_name = "Executive"

                # Replace job grade
                if re.search("E3", job_grade):
                    sheet_name = "Manager"
                    job_grade = "D2"

                if re.search("E4", job_grade):
                    sheet_name = "Senior Managers"
                    job_grade = "M1"

                if re.search("E5", job_grade):
                    sheet_name = "General Managers"
                    job_grade = "C1"

            lc_df = pd.read_excel(LC_file_path, sheet_name=sheet_name, header=1).dropna(axis=1)
            for col in lc_df.columns[1:]:
                lc_df[col] = lc_df[col].astype(str).str.extract("(\d+)")
            # remove double space
            lc_df["Sub-Competency"] = (
                lc_df["Sub-Competency"].apply(lambda x: " ".join(x.split())).str.replace(" : ", ": ")
            )

            lc_list.extend(lc_df[["Sub-Competency"]].values.tolist())
            min_list.extend(lc_df[[job_grade]].values.tolist())
            max_list.extend(lc_df[[job_grade]].values.tolist())

            # if ur in role_level_map.keys():
            #     lc_list.extend(competency_lc_dict[role_level_map[ur]])
            #     min_list.extend(min_lc_dict[role_level_map[ur]])
            #     max_list.extend(max_lc_dict[role_level_map[ur]])
            # else:
            #     lc_list.extend([[""]] * 8)
            #     min_list.extend([[""]] * 8)
            #     max_list.extend([[""]] * 8)

        ProfileItem_Competency_LC.range((13, "E")).value = lc_list
        ProfileItem_Competency_LC.range((13, "K")).value = min_list
        ProfileItem_Competency_LC.range((13, "M")).value = max_list

        # Column H
        source_system_id = [
            [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * 8
            for row in range(12, row_end_talent_profile + 1)
            if wb.sheets[0].range((row, "K")).value != None
        ]
        source_system_id = list(itertools.chain.from_iterable(source_system_id))
        ProfileItem_Competency_LC.range((13, "H")).value = [["={}".format(k)] for k in source_system_id]
        last_row_lc = 13 + len(source_system_id) - 1

        # Column BCDFGIJLN
        for col in "BCDFGIJLN":
            ProfileItem_Competency_LC.range((13, col), (last_row_lc, col)).value = "={}$11".format(col)

        # Column J
        # ProfileItem_Competency_LC.range((13, "J"), (last_row_lc, "J")).value = "Y"

        # Column F
        ProfileItem_Competency_LC.range((13, "F"), (last_row_lc, "F")).value = "'2021/11/01"

        # Column G
        ProfileItem_Competency_LC.range((13, "G"), (last_row_lc, "G")).value = "'4712/12/31"

        # Column O
        z_list = [
            (
                ProfileItem_Competency_LC.range((i, "H"))
                .formula.replace("=", "")
                .replace("TalentProfile!K", "TalentProfile!D"),
                ProfileItem_Competency_LC.range((i, "I")).get_address(False, False),
                ProfileItem_Competency_LC.range((i, "E")).get_address(False, False),
            )
            for i in range(13, last_row_lc + 1)
        ]
        formula_list = [
            [
                '=CONCAT(UPPER(SUBSTITUTE({}," ","_")),"_",{},"_",UPPER(SUBSTITUTE({}," ","_")),"_PI")'.format(
                    k[0], k[1], k[2]
                )
            ]
            for k in z_list
        ]
        ProfileItem_Competency_LC.range((13, "O")).value = formula_list

        # Drop competency with minimum proficiency = 0
        ProfileItem_Competency_LC_df = pd.DataFrame(
            ProfileItem_Competency_LC.range("B13:O{}".format(last_row_lc)).value
        )
        ProfileItem_Competency_LC_df = ProfileItem_Competency_LC_df[
            (ProfileItem_Competency_LC_df[3].notna()) | (ProfileItem_Competency_LC_df[3] == "")
        ]

        ProfileItem_Competency_LC.range((13, "B"), (100000, "O")).clear()
        ProfileItem_Competency_LC.range((13, "B")).value = ProfileItem_Competency_LC_df.values.tolist()
        last_row_lc_2 = (
            ProfileItem_Competency_LC.range("H" + str(ProfileItem_Competency_LC.cells.last_cell.row)).end("up").row
        )
        # Column F
        ProfileItem_Competency_LC.range((13, "F"), (last_row_lc_2, "F")).value = "'2021/11/01"

        # Column G
        ProfileItem_Competency_LC.range((13, "G"), (last_row_lc_2, "G")).value = "'4712/12/31"

        LC_content_item_list = LC_content_item_list.values.tolist()
        for row in range(13, last_row_lc_2 + 1):
            if not ProfileItem_Competency_LC.range((row, "E")).value in LC_content_item_list:
                ProfileItem_Competency_LC.range((row, "E")).color = (255, 0, 0)
            if ProfileItem_Competency_LC.range((row, "L")).value == None:
                ProfileItem_Competency_LC.range((row, "L")).color = (255, 0, 0)
                ProfileItem_Competency_LC.range((row, "N")).color = (255, 0, 0)

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(ProfileItem_Competency_LC.range("B2:O2").value).T
        ProfileItem_Competency_LC_df = pd.DataFrame(
            ProfileItem_Competency_LC.range("B13:O{}".format(last_row_lc_2)).value
        )
        ProfileItem_Competency_LC_df[4] = "2021/11/01"
        ProfileItem_Competency_LC_df[5] = "4712/12/31"

        ProfileItem_Competency_LC_df[[9, 11]] = ProfileItem_Competency_LC_df[[9, 11]].apply(
            pd.to_numeric, downcast="signed"
        )

        ProfileItem_Competency_LC_df = pd.concat([header, ProfileItem_Competency_LC_df])
        ProfileItem_Competency_LC_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ProfileItem-Competency_LC.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    def profile_item_competency_TC(
        wb,
        TC_file_path,
        jcp_file_path,
        position_profile_df,
        position_dat_dir,
        TC_content_item_list,
        read_jcp,
    ):

        ProfileItem_Competency_TC = wb.sheets[21]
        ProfileItem_Competency_TC.range((13, "B"), (500000, "P")).clear()
        technical_competency_df = pd.read_excel(TC_file_path)
        technical_competency_df = technical_competency_df.replace(to_replace=0, value=np.nan).replace(
            "\u200b|\xa0|N/A", "", regex=True
        )
        important_map = {"Core Generic" : 1, "Core Specific" : 2, "Adjacent" : 3}
        technical_competency_df["Important"] = technical_competency_df["Important"].map(important_map)
        # technical_competency_df = technical_competency_df.dropna(
        #     subset=technical_competency_df.iloc[:, 4:].columns, how="all"
        # )
        oracle_column = [
            x
            for x in technical_competency_df.columns
            if re.search("Oracle|Compentecy Technical|ContentItem", x, flags=re.I)
        ][0]
        technical_competency_df[oracle_column] = (
            technical_competency_df[oracle_column].apply(lambda x: " ".join(str(x).split())).str.replace("–", "-")
        )
        spur_id_tc_occ = technical_competency_df["SPUR ID"].value_counts().sort_index()

        if read_jcp == True:
            jcp_df = pd.read_excel(jcp_file_path).replace("\u200b|\xa0|N\A", "", regex=True)
            jcp_df["competency"] = jcp_df["Ti Name"].str.strip() + " " + jcp_df["Ti Number"].str.strip()
            jcp_df["competency"] = (
                jcp_df["competency"]
                .apply(lambda x: " ".join(x.split()) if isinstance(x, str) else x)
                .str.replace("–", "-")
            )
            jcp_df["Position Id"] = jcp_df["Position Id"].astype(str).str.zfill(8)
            important_map = {"Core Generic" : 1, "Core Specific" : 2, "Adjacent" : 3}
            jcp_df["Category"] = jcp_df["Category"].map(important_map)
            #important_map = {"Core Generic" : 1, "Core Specific" : 2, "Adjacent" : 3}
            #technical_competency_df["Important"] = technical_competency_df["Important"].map(important_map)
        else:
            jcp_df = pd.DataFrame([np.nan], columns=["Position Id"])

        tc_ref_list = []
        tc_name_list = []
        min_tc_list = []
        max_tc_list = []
        ProfileItem_Competency_TC_Important_list = []

        base_columns = ["SPUR ID", "TI Name", "TI Number", oracle_column]
        row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
        for row in range(12, row_end_talent_profile + 1):
            if wb.sheets[0].range((row, "K")).value == None:
                break

            ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
            position_id = str(re.search("[^_]+$", wb.sheets[0].range((row, "D")).value).group()).zfill(8)
            role_level_column = [
                x for x in position_profile_df.columns if re.search("role level|role$", x.lower().strip())
            ][0]
            role_level = position_profile_df[position_profile_df["Pos ID"] == position_id][role_level_column].values[0]
            jg_column = [x for x in position_profile_df.columns if "jg" in x.lower()][0]
            conso_JG = position_profile_df[position_profile_df["Pos ID"] == position_id][jg_column].values[0]
            # if not isinstance(conso_JG, str):
            # conso_JG = WS_df[WS_df["Pos ID (as per ZHPLA 3/5/2021)"] == position_id]["Conso JG"].iloc[0]
            conso_JG = re.sub(r"Est.|Eqv.", r"", conso_JG).strip()
            if (
                ur_id in technical_competency_df["SPUR ID"].values.tolist()
                or position_id in jcp_df["Position Id"].values.tolist()
            ):
                if read_jcp == True:
                    if position_id in jcp_df["Position Id"].values.tolist():
                        pid_df = jcp_df[jcp_df["Position Id"] == position_id]
                        tc_ref_list.extend(
                            [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * len(pid_df)
                        )
                        ProfileItem_Competency_TC_Important_list.extend(
                            jcp_df[jcp_df["Position Id"] == position_id][["Category"]].values.tolist()
                        )
                    elif ur_id in technical_competency_df["SPUR ID"].values.tolist():
                        tc_ref_list.extend(
                            [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_tc_occ[ur_id]
                        )
                        ProfileItem_Competency_TC_Important_list.extend(
                            technical_competency_df[technical_competency_df["SPUR ID"] == ur_id][["Important"]].values.tolist()
                    )
                    else:
                        continue
                else:
                    tc_ref_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_tc_occ[ur_id]
                    )
                    ProfileItem_Competency_TC_Important_list.extend(
                         technical_competency_df[technical_competency_df["SPUR ID"] == ur_id][["Important"]].values.tolist()
                     )

            else:
                if run_all == True:
                    tc_ref_list.extend([wb.sheets[0].range((row, "K")).get_address(True, False, True)])
                else:
                    continue

            if (
                ur_id in technical_competency_df["SPUR ID"].values.tolist()
                or position_id in jcp_df["Position Id"].values.tolist()
            ):
                if read_jcp == True:
                    if position_id in jcp_df["Position Id"].values.tolist():
                        tc_name_list.extend(pid_df[["competency"]].values.tolist())
                        min_tc_list.extend(pid_df[["Proficiency Level"]].values.tolist())
                        max_tc_list.extend(pid_df[["Proficiency Level"]].values.tolist())
                        #ProfileItem_Competency_TC_Important_list.extend(pid_df[["Category"]].values.tolist())
                        #ProfileItem_Competency_TC_Important_list.extend(pid_df[["Important"]].values.tolist())
                    elif ur_id in technical_competency_df["SPUR ID"].values.tolist():
                        tc_name_list.extend(
                            technical_competency_df[technical_competency_df["SPUR ID"] == ur_id][
                                [oracle_column]
                            ].values.tolist()
                        )
                        #ProfileItem_Competency_TC_Important_list.extend(
                            #technical_competency_df[technical_competency_df["SPUR ID"] == ur_id][["Important"]].values.tolist()
                    #)

                        temp_df = technical_competency_df[technical_competency_df["SPUR ID"] == ur_id]

                        if role_level in ["Staff", "Principal", "Custodian"]:
                            if temp_df[role_level].isnull().all():
                                if role_level == "Staff":
                                    min_tc_list.extend([[x] for x in temp_df["D2"].values.tolist()])
                                    max_tc_list.extend([[x] for x in temp_df["D2"].values.tolist()])
                                elif role_level == "Principal":
                                    min_tc_list.extend([[x] for x in temp_df["M1"].values.tolist()])
                                    max_tc_list.extend([[x] for x in temp_df["M1"].values.tolist()])
                                else:
                                    min_tc_list.extend([[x] for x in temp_df["C1"].values.tolist()])
                                    max_tc_list.extend([[x] for x in temp_df["C1"].values.tolist()])
                            else:
                                min_tc_list.extend([[x] for x in temp_df[role_level].values.tolist()])
                                max_tc_list.extend([[x] for x in temp_df[role_level].values.tolist()])

                        else:
                            if conso_JG in temp_df.columns:
                                min_tc_list.extend([[x] for x in temp_df[conso_JG].values.tolist()])
                                max_tc_list.extend([[x] for x in temp_df[conso_JG].values.tolist()])
                            elif conso_JG == "E3":
                                min_tc_list.extend([[x] for x in temp_df["D2"].values.tolist()])
                                max_tc_list.extend([[x] for x in temp_df["D2"].values.tolist()])
                            elif conso_JG == "E4":
                                min_tc_list.extend([[x] for x in temp_df["M1"].values.tolist()])
                                max_tc_list.extend([[x] for x in temp_df["M1"].values.tolist()])
                            else:
                                min_tc_list.extend([[x] for x in temp_df["C1"].values.tolist()])
                                max_tc_list.extend([[x] for x in temp_df["C1"].values.tolist()])
                else:
                    tc_name_list.extend(
                        technical_competency_df[technical_competency_df["SPUR ID"] == ur_id][
                            [oracle_column]
                        ].values.tolist()
                    )

                    temp_df = technical_competency_df[technical_competency_df["SPUR ID"] == ur_id]
                    # rating_cols = [
                    #     "A1",
                    #     "A2",
                    #     "A3",
                    #     "D1",
                    #     "D2",
                    #     "D3",
                    #     "M1",
                    #     "M2",
                    #     "C1",
                    #     "C2",
                    #     "Staff",
                    #     "Principal",
                    #     "Custodian",
                    #     "H1",
                    #     "H2",
                    # ]
                    # temp_df = temp_df.loc[:, "A1":]
                    # temp_df = temp_df.applymap(
                    #     lambda x: int(float(re.sub("\s+|\u200b", "", str(x)))) if str(x) != "nan" else x
                    # )

                    # if role_level in ["Executive", "Head", "Manager", "Chief"]:
                    #     if conso_JG in temp_df.columns:
                    #         # temp_df = temp_df.loc[:, conso_JG]
                    #         min_tc_list.extend([[x] for x in temp_df[conso_JG].values.tolist()])
                    #         max_tc_list.extend([[x] for x in temp_df[conso_JG].values.tolist()])

                    #     else:
                    #         pass

                    if role_level in ["Staff", "Principal", "Custodian"]:
                        # if role_level in temp_df.columns:
                        # temp_df = temp_df.loc[:, role_level]
                        if temp_df[role_level].isnull().all():
                            if role_level == "Staff":
                                min_tc_list.extend([[x] for x in temp_df["D2"].values.tolist()])
                                max_tc_list.extend([[x] for x in temp_df["D2"].values.tolist()])
                            elif role_level == "Principal":
                                min_tc_list.extend([[x] for x in temp_df["M1"].values.tolist()])
                                max_tc_list.extend([[x] for x in temp_df["M1"].values.tolist()])
                            else:
                                min_tc_list.extend([[x] for x in temp_df["C1"].values.tolist()])
                                max_tc_list.extend([[x] for x in temp_df["C1"].values.tolist()])
                        else:
                            min_tc_list.extend([[x] for x in temp_df[role_level].values.tolist()])
                            max_tc_list.extend([[x] for x in temp_df[role_level].values.tolist()])

                    else:
                        if conso_JG in temp_df.columns:
                            min_tc_list.extend([[x] for x in temp_df[conso_JG].values.tolist()])
                            max_tc_list.extend([[x] for x in temp_df[conso_JG].values.tolist()])
                        elif conso_JG == "E3":
                            min_tc_list.extend([[x] for x in temp_df["D2"].values.tolist()])
                            max_tc_list.extend([[x] for x in temp_df["D2"].values.tolist()])
                        elif conso_JG == "E4":
                            min_tc_list.extend([[x] for x in temp_df["M1"].values.tolist()])
                            max_tc_list.extend([[x] for x in temp_df["M1"].values.tolist()])
                        else:
                            min_tc_list.extend([[x] for x in temp_df["C1"].values.tolist()])
                            max_tc_list.extend([[x] for x in temp_df["C1"].values.tolist()])

                    # temp_df = temp_df.fillna(0)
                    # if position_id == 2197970:
                    #     print(temp_df)

                    # min_tc_list.extend([[int(x)] if not np.isnan(x) else [x] for x in temp_df.min(axis=1).tolist()])
                    # max_tc_list.extend([[int(x)] if not np.isnan(x) else [x] for x in temp_df.max(axis=1).tolist()])

            else:
                if run_all == True:
                    tc_name_list.extend([[""]])
                    min_tc_list.extend([[""]])
                    max_tc_list.extend([[""]])
                    ProfileItem_Competency_TC_Important_list.extend([[""]])
                else:
                    continue

        # Column H
        ProfileItem_Competency_TC.range((13, "H")).value = [["={}".format(k)] for k in tc_ref_list]
        row_end_competency_TC = (
            ProfileItem_Competency_TC.range("H" + str(ProfileItem_Competency_TC.cells.last_cell.row)).end("up").row
        )

        # Column E
        ProfileItem_Competency_TC.range((13, "E")).value = tc_name_list

        # Column L
        ProfileItem_Competency_TC.range((13, "L")).value = min_tc_list

        # Column N
        ProfileItem_Competency_TC.range((13, "N")).value = max_tc_list

        # Column BCDFGIKMO
        for col in "BCDFGIKMO":
            ProfileItem_Competency_TC.range((13, col), (row_end_competency_TC, col)).value = "={}$11".format(col)

        # Column J
        ProfileItem_Competency_TC.range((13, "J")).value = ProfileItem_Competency_TC_Important_list

        # Column K
        # ProfileItem_Competency_TC.range((13, "J"), (row_end_competency_TC, "J")).value = "Y"

        # Column F
        ProfileItem_Competency_TC.range((13, "F"), (row_end_competency_TC, "F")).value = "'2021/11/01"

        # Column G
        ProfileItem_Competency_TC.range((13, "G"), (row_end_competency_TC, "G")).value = "'4712/12/31"

        # Column P
        z_list = [
            (
                ProfileItem_Competency_TC.range((i, "H"))
                .formula.replace("=", "")
                .replace("TalentProfile!K", "TalentProfile!D"),
                ProfileItem_Competency_TC.range((i, "I")).get_address(False, False),
                ProfileItem_Competency_TC.range((i, "E")).get_address(False, False),
            )
            for i in range(13, row_end_competency_TC + 1)
        ]
        formula_list = [
            [
                '=CONCAT(UPPER(SUBSTITUTE({}," ","_")),"_",{},"_",UPPER(SUBSTITUTE({}," ","_")),"_PI")'.format(
                    k[0], k[1], k[2]
                )
            ]
            for k in z_list
        ]
        ProfileItem_Competency_TC.range((13, "P")).value = formula_list

        # Drop competency with minimum proficiency = 0
        ProfileItem_Competency_TC_df = pd.DataFrame(
            ProfileItem_Competency_TC.range("B13:P{}".format(row_end_competency_TC)).value
        )
        ProfileItem_Competency_TC_df = ProfileItem_Competency_TC_df[ProfileItem_Competency_TC_df[10].notna()]
        ProfileItem_Competency_TC_df = ProfileItem_Competency_TC_df[
            (ProfileItem_Competency_TC_df[3].notna()) | (ProfileItem_Competency_TC_df[3] == "")
        ]
        ProfileItem_Competency_TC_df[4] = "2021/11/01"
        ProfileItem_Competency_TC_df[5] = "4712/12/31"

        # ProfileItem_Competency_TC_df[[10, 12]] = ProfileItem_Competency_TC_df[[10, 12]].apply(
        #     pd.to_numeric, downcast="signed"
        # )
        ProfileItem_Competency_TC_df[8] = ProfileItem_Competency_TC_df[8].fillna(-1).astype(int).replace(-1, "")
        ProfileItem_Competency_TC_df[10] = ProfileItem_Competency_TC_df[10].fillna(-1).astype(int).replace(-1, "")
        ProfileItem_Competency_TC_df[12] = ProfileItem_Competency_TC_df[12].fillna(-1).astype(int).replace(-1, "")
        ProfileItem_Competency_TC_df = ProfileItem_Competency_TC_df.drop_duplicates(subset=[14])
        # ProfileItem_Competency_TC_df = ProfileItem_Competency_TC_df[~ProfileItem_Competency_TC_df[15].duplicated()]

        ProfileItem_Competency_TC.range((13, "B"), (500000, "P")).clear()
        ProfileItem_Competency_TC.range((13, "B")).value = ProfileItem_Competency_TC_df.values.tolist()
        row_end_competency_TC_2 = (
            ProfileItem_Competency_TC.range("H" + str(ProfileItem_Competency_TC.cells.last_cell.row)).end("up").row
        )
        # Column F
        ProfileItem_Competency_TC.range((13, "F"), (row_end_competency_TC_2, "F")).value = "'2021/11/01"

        # Column G
        ProfileItem_Competency_TC.range((13, "G"), (row_end_competency_TC_2, "G")).value = "'4712/12/31"

        # ProfileItem_Competency_TC_df[[10, 12]] = ProfileItem_Competency_TC_df[[10, 12]].apply(
        #     pd.to_numeric, downcast="signed"
        # )

        # TC_content_item_list = TC_content_item_list.values.tolist()
        # for row in range(35, row_end_competency_TC_2 + 1):
        #     if not ProfileItem_Competency_TC.range((row, "E")).value in TC_content_item_list:
        #         ProfileItem_Competency_TC.range((row, "E")).color = (255, 0, 0)
        #     if ProfileItem_Competency_TC.range((row, "L")).value == None:
        #         ProfileItem_Competency_TC.range((row, "L")).color = (255, 0, 0)
        #         ProfileItem_Competency_TC.range((row, "N")).color = (255, 0, 0)

        row_end_profile_relation = wb.sheets[1].range("H" + str(wb.sheets[1].cells.last_cell.row)).end("up").row
        pid_list = [str(x).zfill(8) for x in wb.sheets[1].range((11, "H"), (row_end_profile_relation, "H")).value]
        source_system_id_list = list(set(ProfileItem_Competency_TC_df[6].values.tolist()))

        pid_TC_missing_list = []
        for pid in pid_list:
            if all([not re.search(str(pid), x) for x in source_system_id_list]):
                pid_TC_missing_list.append(str(pid))
        if len(pid_TC_missing_list) != 0:
            logging.warning("[Data validation] {} PID have no TC data".format(len(pid_TC_missing_list)))
            with open(log_dir + "\\" + f"{skg_name}_PID_without_TC.txt", "w") as f:
                f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(pid_TC_missing_list, start=1)))
        else:
            if f"{skg_name}_PID_without_TC.txt" in log_dir_list:
                os.remove(log_dir + "\\" + f"{skg_name}_PID_without_TC.txt")

        header = pd.DataFrame(ProfileItem_Competency_TC.range("B2:P2").value).T
        ProfileItem_Competency_TC_df = pd.concat([header, ProfileItem_Competency_TC_df])
        ProfileItem_Competency_TC_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ProfileItem-Competency_TC.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    def profile_item_degree(wb, degree_df, position_dat_dir, AreaOfStudy_content_item_list):
        ProfileItem_Degree = wb.sheets[8]
        ProfileItem_Degree.range((12, "B"), (10000, "R")).clear()

        spur_id_degree_occ = degree_df["SPUR ID"].value_counts().sort_index()

        degree_id_list = []
        degree_importance_list = []
        edu_level_list = []
        degree_name_list = []

        row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
        for row in range(12, row_end_talent_profile + 1):
            if wb.sheets[0].range((row, "K")).value == None:
                break

            ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
            ur_id_degree_df = degree_df[degree_df["SPUR ID"] == ur_id]
            position_id = str(re.search("[^_]+$", wb.sheets[0].range((row, "D")).value).group()).zfill(8)
            role_level_column = [
                x for x in position_profile_df.columns if re.search("role level|role$", x.lower().strip())
            ][0]
            role_level = position_profile_df[position_profile_df["Pos ID"] == position_id][role_level_column].values[0]
            jg_column = [x for x in position_profile_df.columns if "jg" in x.lower()][0]
            conso_JG = position_profile_df[position_profile_df["Pos ID"] == position_id][jg_column].values[0]
            # conso_JG = WS_df[WS_df["Pos ID (as per ZHPLA 3/5/2021)"] == position_id]["Conso JG"].iloc[0]
            conso_JG = re.sub(r"Est.|Eqv.", r"", conso_JG).strip()

            if ur_id in spur_id_degree_occ.index:
                if all(ur_id_degree_df["JG"].notna()) and any([str(conso_JG) in x for x in ur_id_degree_df["JG"]]):
                    ur_id_degree_df = ur_id_degree_df.iloc[[conso_JG in x for x in ur_id_degree_df["JG"]]]
                    degree_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * len(ur_id_degree_df)
                    )

                else:
                    degree_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_degree_occ[ur_id]
                    )

            else:
                if run_all == True:
                    degree_id_list.extend([wb.sheets[0].range((row, "K")).get_address(True, False, True)])
                else:
                    continue

            if ur_id in degree_df["SPUR ID"].values.tolist():
                if all(ur_id_degree_df["JG"].notna()) and any([str(conso_JG) in x for x in ur_id_degree_df["JG"]]):
                    degree_importance_list.extend(
                        ur_id_degree_df[ur_id_degree_df["SPUR ID"] == ur_id][["Importance"]].values.tolist()
                    )
                    edu_level_list.extend(
                        ur_id_degree_df[ur_id_degree_df["SPUR ID"] == ur_id][[degree_column]].values.tolist()
                    )
                    degree_name_list.extend(
                        ur_id_degree_df[ur_id_degree_df["SPUR ID"] == ur_id][[area_of_study_column]].values.tolist()
                    )
                else:
                    degree_importance_list.extend(
                        degree_df[degree_df["SPUR ID"] == ur_id][["Importance"]].values.tolist()
                    )
                    edu_level_list.extend(degree_df[degree_df["SPUR ID"] == ur_id][[degree_column]].values.tolist())
                    degree_name_list.extend(
                        degree_df[degree_df["SPUR ID"] == ur_id][[area_of_study_column]].values.tolist()
                    )
            else:
                if run_all == True:
                    degree_importance_list.extend([[""]])
                    edu_level_list.extend([[""]])
                    degree_name_list.extend([[""]])
                else:
                    continue

        # Column H
        ProfileItem_Degree.range((12, "H")).value = [["={}".format(k)] for k in degree_id_list]

        # Column K
        # ProfileItem_Degree.range((12, "K")).value = degree_importance_list

        # Column E
        ProfileItem_Degree.range((12, "E")).value = edu_level_list

        # Column P
        ProfileItem_Degree.range((12, "P")).value = degree_name_list

        row_end_degree_sheet = (
            ProfileItem_Degree.range("H" + str(ProfileItem_Degree.cells.last_cell.row)).end("up").row
        )
        # Column BCDFGIMNOQ
        for k in "BCDFGIMNOQ":
            ProfileItem_Degree.range("{}12:{}{}".format(k, k, row_end_degree_sheet)).value = "={}$11".format(k)

        # Column F
        ProfileItem_Degree.range((12, "F"), (row_end_degree_sheet, "F")).value = "'2021/11/01"

        # Column G
        ProfileItem_Degree.range((12, "G"), (row_end_degree_sheet, "G")).value = "'4712/12/31"

        # Column R
        z_list = [
            (
                ProfileItem_Degree.range((i, "H"))
                .formula.replace("=", "")
                .replace("TalentProfile!K", "TalentProfile!D"),
                ProfileItem_Degree.range((i, "I")).get_address(False, False),
                ProfileItem_Degree.range((i, "E")).get_address(False, False),
                ProfileItem_Degree.range((i, "P")).get_address(False, False),
            )
            for i in range(12, row_end_degree_sheet + 1)
        ]
        formula_list = [
            [
                '=CONCAT(UPPER(SUBSTITUTE({}," ","_")),"_",{},"_",UPPER(SUBSTITUTE({}," ","_")),"_",UPPER(SUBSTITUTE({}," ","_")),"_PI")'.format(
                    k[0], k[1], k[2], k[3]
                )
            ]
            for k in z_list
        ]
        ProfileItem_Degree.range((12, "R")).value = formula_list

        AreaOfStudy_content_item_list = AreaOfStudy_content_item_list.values.tolist()
        AreaOfStudy_content_item_list = [x.strip() for x in AreaOfStudy_content_item_list]
        for row in range(12, row_end_degree_sheet + 1):
            if not ProfileItem_Degree.range((row, "E")).value in AreaOfStudy_content_item_list:
                ProfileItem_Degree.range((row, "E")).color = (255, 0, 0)
            if ProfileItem_Degree.range((row, "K")).value == None:
                ProfileItem_Degree.range((row, "K")).color = (255, 0, 0)

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(ProfileItem_Degree.range("B2:R2").value).T
        ProfileItem_Degree_df = pd.DataFrame(ProfileItem_Degree.range("B12:R{}".format(row_end_degree_sheet)).value)
        ProfileItem_Degree_df[4] = "2021/11/01"
        ProfileItem_Degree_df[5] = "4712/12/31"
        # ProfileItem_Degree_df[9] = ProfileItem_Degree_df[9].astype(int)

        ProfileItem_Degree.range((12, "B"), (10000, "R")).clear()
        ProfileItem_Degree.range((12, "B")).value = ProfileItem_Degree_df.values.tolist()

        ProfileItem_Degree.range((12, "F"), (row_end_degree_sheet, "F")).value = "'2021/11/01"
        ProfileItem_Degree.range((12, "G"), (row_end_degree_sheet, "G")).value = "'4712/12/31"

        ProfileItem_Degree_df = pd.concat([header, ProfileItem_Degree_df])
        ProfileItem_Degree_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ProfileItem-Degree.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    def profile_item_language(wb, position_dat_dir):
        ProfileItem_Language = wb.sheets[12]
        ProfileItem_Language.range((12, "B"), (10000, "R")).clear()

        row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
        h12_language = [
            wb.sheets[0].range((row, "K")).get_address(True, False, True)
            for row in range(12, row_end_talent_profile + 1)
            if wb.sheets[0].range((row, "K")).value != None
        ]
        h12_language = [["={}".format(k)] for k in h12_language]

        ProfileItem_Language.range((12, "H")).value = h12_language

        row_end_language_sheet = (
            ProfileItem_Language.range("H" + str(ProfileItem_Language.cells.last_cell.row)).end("up").row
        )
        # Column BCDFGIJKLMNOPQ
        for k in "BCDFGIJKLMNOPQ":
            ProfileItem_Language.range("{}12:{}{}".format(k, k, row_end_language_sheet)).value = "={}$11".format(k)

        # Column E
        ProfileItem_Language.range((12, "E"), (row_end_language_sheet, "E")).value = "English"

        # Column F
        ProfileItem_Language.range((12, "F"), (row_end_language_sheet, "F")).value = "'2021/11/01"

        # Column G
        ProfileItem_Language.range((12, "G"), (row_end_language_sheet, "G")).value = "'4712/12/31"

        # Column R
        z_list = [
            (
                ProfileItem_Language.range((i, "H"))
                .formula.replace("=", "")
                .replace("TalentProfile!K", "TalentProfile!D"),
                ProfileItem_Language.range((i, "I")).get_address(False, False),
                ProfileItem_Language.range((i, "E")).get_address(False, False),
            )
            for i in range(12, row_end_language_sheet + 1)
        ]
        formula_list = [
            ['=CONCAT(UPPER(SUBSTITUTE({}," ","_")),"_",UPPER({}),"_",UPPER({}),"_PI")'.format(k[0], k[1], k[2])]
            for k in z_list
        ]
        ProfileItem_Language.range((12, "R")).value = formula_list

        # Drop competency with minimum proficiency = 0
        header = pd.DataFrame(ProfileItem_Language.range("B2:R2").value).T
        ProfileItem_Language_df = pd.DataFrame(
            ProfileItem_Language.range("B12:R{}".format(row_end_language_sheet)).value
        )
        ProfileItem_Language_df[4] = "2021/11/01"
        ProfileItem_Language_df[5] = "4712/12/31"
        # ProfileItem_Language_df[15] = ProfileItem_Language_df[15].astype(int)

        ProfileItem_Language.range((12, "B"), (10000, "R")).clear()
        ProfileItem_Language.range((12, "B")).value = ProfileItem_Language_df.values.tolist()

        ProfileItem_Language.range((12, "F"), (row_end_language_sheet, "F")).value = "'2021/11/01"
        ProfileItem_Language.range((12, "G"), (row_end_language_sheet, "G")).value = "'4712/12/31"

        ProfileItem_Language_df = pd.concat([header, ProfileItem_Language_df])
        ProfileItem_Language_df.to_csv(
            position_dat_dir + "\\" + f"{skg_name}_ProfileItem-Language.dat",
            header=None,
            index=None,
            sep="|",
            mode="w",
            date_format="%Y/%m/%d",
        )

    def profile_item_membership(
        wb,
        membership_df,
        position_profile_df,
        position_dat_dir,
        membership_content_item_list,
    ):
        ProfileItem_Membership = wb.sheets[14]
        ProfileItem_Membership.range((12, "B"), (10000, "N")).clear()

        spur_id_membership_occ = membership_df["SPUR ID"].value_counts().sort_index()

        talent_profile_list = []
        membership_list = []
        membership_importance_list = []

        row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
        for row in range(12, row_end_talent_profile + 1):
            if wb.sheets[0].range((row, "K")).value == None:
                break

            ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
            ur_id_membership_df = membership_df[membership_df["SPUR ID"] == ur_id]
            position_id = str(re.search("[^_]+$", wb.sheets[0].range((row, "D")).value).group()).zfill(8)
            # print(position_id)
            role_level_column = [
                x for x in position_profile_df.columns if re.search("role level|role$", x.lower().strip())
            ][0]
            # print(role_level_column)
            # role_level_column = 'Role Level'
            role_level = position_profile_df[position_profile_df["Pos ID"] == position_id][role_level_column].values[0]
            jg_column = [x for x in position_profile_df.columns if "jg" in x.lower()][0]
            conso_JG = position_profile_df[position_profile_df["Pos ID"] == position_id][jg_column].values[0]
            # conso_JG = WS_df[WS_df["Pos ID (as per ZHPLA 3/5/2021)"] == position_id]["Conso JG"].iloc[0]
            conso_JG = re.sub(r"Est.|Eqv.", r"", conso_JG).strip()

            if ur_id in spur_id_membership_occ.index:
                if all(ur_id_membership_df["JG"].notna()) and any(
                    [str(conso_JG) in x for x in ur_id_membership_df["JG"]]
                ):
                    ur_id_membership_df = ur_id_membership_df.iloc[[conso_JG in x for x in ur_id_membership_df["JG"]]]
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * len(ur_id_membership_df)
                    )

                else:
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_membership_occ[ur_id]
                    )

            else:
                continue

            if ur_id in membership_df["SPUR ID"].values.tolist():
                if all(ur_id_membership_df["JG"].notna()) and any(
                    [str(conso_JG) in x for x in ur_id_membership_df["JG"]]
                ):
                    membership_list.extend(
                        ur_id_membership_df[ur_id_membership_df["SPUR ID"] == ur_id][
                            [membership_column]
                        ].values.tolist()
                    )
                    membership_importance_list.extend(
                        ur_id_membership_df[ur_id_membership_df["SPUR ID"] == ur_id][["Importance"]].values.tolist()
                    )

                else:
                    membership_list.extend(
                        membership_df[membership_df["SPUR ID"] == ur_id][[membership_column]].values.tolist()
                    )
                    membership_importance_list.extend(
                        degree_df[degree_df["SPUR ID"] == ur_id][["Importance"]].values.tolist()
                    )

            else:
                continue

        if membership_list != []:
            # Column H
            ProfileItem_Membership.range((12, "H")).value = [["={}".format(k)] for k in talent_profile_list]

            # Column E
            ProfileItem_Membership.range((12, "E")).value = membership_list

            row_end_membership_sheet = (
                ProfileItem_Membership.range("H" + str(ProfileItem_Membership.cells.last_cell.row)).end("up").row
            )
            # Column BCDFGIJKLM
            for k in "BCDIJKLM":
                ProfileItem_Membership.range(
                    "{}12:{}{}".format(k, k, row_end_membership_sheet)
                ).value = "={}$11".format(k)

            # Column F
            ProfileItem_Membership.range((12, "F"), (row_end_membership_sheet, "F")).value = "'2021/11/01"

            # Column G
            ProfileItem_Membership.range((12, "G"), (row_end_membership_sheet, "G")).value = "'4712/12/31"

            # Column N
            z_list = [
                (
                    ProfileItem_Membership.range((i, "H"))
                    .formula.replace("=", "")
                    .replace("TalentProfile!K", "TalentProfile!D"),
                    ProfileItem_Membership.range((i, "D")).get_address(False, False),
                    ProfileItem_Membership.range((i, "E")).get_address(False, False),
                )
                for i in range(12, row_end_membership_sheet + 1)
            ]
            formula_list = [
                [
                    '=CONCAT(UPPER(SUBSTITUTE({}," ","_")),"_",{},"_",UPPER(SUBSTITUTE({}," ","_")),"_PI")'.format(
                        k[0], k[1], k[2]
                    )
                ]
                for k in z_list
            ]
            ProfileItem_Membership.range((12, "N")).value = formula_list

            membership_content_item_list = membership_content_item_list.values.tolist()
            for row in range(12, row_end_membership_sheet + 1):
                if not ProfileItem_Membership.range((row, "E")).value in membership_content_item_list:
                    ProfileItem_Membership.range((row, "E")).color = (255, 0, 0)

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ProfileItem_Membership.range("B2:N2").value).T
            ProfileItem_Membership_df = pd.DataFrame(
                ProfileItem_Membership.range("B12:N{}".format(row_end_membership_sheet)).value
            )
            ProfileItem_Membership_df[4] = "2021/11/01"
            ProfileItem_Membership_df[5] = "4712/12/31"
            # ProfileItem_Membership_df[10] = ProfileItem_Membership_df[10].astype(int)
            # ProfileItem_Membership_df = ProfileItem_Membership_df.drop_duplicates(subset=[13], keep="first")

            ProfileItem_Membership.range((12, "B"), (10000, "N")).clear()
            ProfileItem_Membership.range((12, "B")).value = ProfileItem_Membership_df.values.tolist()
            row_end_membership_sheet_2 = (
                ProfileItem_Membership.range("H" + str(ProfileItem_Membership.cells.last_cell.row)).end("up").row
            )

            ProfileItem_Membership.range((12, "F"), (row_end_membership_sheet_2, "F")).value = "'2021/11/01"
            ProfileItem_Membership.range((12, "G"), (row_end_membership_sheet_2, "G")).value = "'4712/12/31"

            ProfileItem_Membership_df = pd.concat([header, ProfileItem_Membership_df])
            ProfileItem_Membership_df.to_csv(
                position_dat_dir + "\\" + f"{skg_name}_ProfileItem-Membership.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
                date_format="%Y/%m/%d",
            )

        else:
            pass

    def profile_item_awards(wb, awards_df, position_profile_df, position_dat_dir, awards_content_item_list):
        ProfileItem_Awards = wb.sheets[10]
        ProfileItem_Awards.range((12, "B"), (10000, "O")).clear()

        spur_id_awards_occ = awards_df["SPUR ID"].value_counts().sort_index()

        talent_profile_list = []
        awards_list = []
        importance_list = []

        row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
        for row in range(12, row_end_talent_profile + 1):
            if wb.sheets[0].range((row, "K")).value == None:
                break

            ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
            ur_id_awards_df = awards_df[awards_df["SPUR ID"] == ur_id]
            position_id = str(re.search("[^_]+$", wb.sheets[0].range((row, "D")).value).group()).zfill(8)
            role_level_column = [
                x for x in position_profile_df.columns if re.search("role level|role$", x.lower().strip())
            ][0]
            role_level = position_profile_df[position_profile_df["Pos ID"] == position_id][role_level_column].values[0]
            jg_column = [x for x in position_profile_df.columns if "jg" in x.lower()][0]
            conso_JG = position_profile_df[position_profile_df["Pos ID"] == position_id][jg_column].values[0]
            # conso_JG = WS_df[WS_df["Pos ID (as per ZHPLA 3/5/2021)"] == position_id]["Conso JG"].iloc[0]
            conso_JG = re.sub(r"Est.|Eqv.", r"", conso_JG).strip()

            if ur_id in spur_id_awards_occ.index:
                if all(ur_id_awards_df["JG"].notna()) and any([str(conso_JG) in x for x in ur_id_awards_df["JG"]]):
                    ur_id_awards_df = ur_id_awards_df.iloc[[conso_JG in x for x in ur_id_awards_df["JG"]]]
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * len(ur_id_awards_df)
                    )

                else:
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_awards_occ[ur_id]
                    )

            else:
                continue

            if ur_id in awards_df["SPUR ID"].values.tolist():
                if all(ur_id_awards_df["JG"].notna()) and any([str(conso_JG) in x for x in ur_id_awards_df["JG"]]):
                    awards_list.extend(
                        ur_id_awards_df[ur_id_awards_df["SPUR ID"] == ur_id][[membership_column]].values.tolist()
                    )
                    importance_list.extend(
                        ur_id_awards_df[ur_id_awards_df["SPUR ID"] == ur_id][["Importance"]].values.tolist()
                    )

                else:
                    awards_list.extend(awards_df[awards_df["SPUR ID"] == ur_id][[awards_column]].values.tolist())
                    importance_list.extend(degree_df[degree_df["SPUR ID"] == ur_id][["Importance"]].values.tolist())

            else:
                continue

        if awards_list != []:
            # Column H
            ProfileItem_Awards.range((12, "H")).value = [["={}".format(k)] for k in talent_profile_list]

            # Column E
            ProfileItem_Awards.range((12, "E")).value = awards_list

            # Column J
            ProfileItem_Awards.range((12, "J")).value = importance_list

            row_end_awards_sheet = ProfileItem_Awards.range("H" + str(wb.sheets[10].cells.last_cell.row)).end("up").row
            # Column BCDFGIKMN
            for k in "BCDFGIKMN":
                ProfileItem_Awards.range("{}12:{}{}".format(k, k, row_end_awards_sheet)).value = "={}$11".format(k)

            # Column F
            ProfileItem_Awards.range((12, "F"), (row_end_awards_sheet, "F")).value = "'2021/11/01"

            # Column G
            ProfileItem_Awards.range((12, "G"), (row_end_awards_sheet, "G")).value = "'4712/12/31"

            # Column N
            z_list = [
                (
                    ProfileItem_Awards.range((i, "H"))
                    .formula.replace("=", "")
                    .replace("TalentProfile!K", "TalentProfile!D"),
                    ProfileItem_Awards.range((i, "D")).get_address(False, False),
                    ProfileItem_Awards.range((i, "E")).get_address(False, False),
                )
                for i in range(12, row_end_awards_sheet + 1)
            ]
            formula_list = [
                [
                    '=CONCAT(UPPER(SUBSTITUTE({}," ","_")),"_",{},"_",UPPER(SUBSTITUTE({}," ","_")),"_PI")'.format(
                        k[0], k[1], k[2]
                    )
                ]
                for k in z_list
            ]
            ProfileItem_Awards.range((12, "O")).value = formula_list

            awards_content_item_list = awards_content_item_list.values.tolist()
            for row in range(12, row_end_awards_sheet + 1):
                if not ProfileItem_Awards.range((row, "E")).value in awards_content_item_list:
                    ProfileItem_Awards.range((row, "E")).color = (255, 0, 0)

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ProfileItem_Awards.range("B2:O2").value).T
            ProfileItem_Awards_df = pd.DataFrame(
                ProfileItem_Awards.range("B12:O{}".format(row_end_awards_sheet)).value
            )
            ProfileItem_Awards_df[4] = "2021/11/01"
            ProfileItem_Awards_df[5] = "4712/12/31"
            ProfileItem_Awards_df[8] = ProfileItem_Awards_df[8].astype(int)
            # ProfileItem_Awards_df = ProfileItem_Awards_df.drop_duplicates(subset=[12], keep="first")

            ProfileItem_Awards.range((12, "B"), (10000, "O")).clear()
            ProfileItem_Awards.range((12, "B")).value = ProfileItem_Awards_df.values.tolist()
            row_end_awards_sheet_2 = (
                ProfileItem_Awards.range("H" + str(ProfileItem_Awards.cells.last_cell.row)).end("up").row
            )

            ProfileItem_Awards.range((12, "F"), (row_end_awards_sheet_2, "F")).value = "'2021/11/01"
            ProfileItem_Awards.range((12, "G"), (row_end_awards_sheet_2, "G")).value = "'4712/12/31"

            ProfileItem_Awards_df = pd.concat([header, ProfileItem_Awards_df])
            ProfileItem_Awards_df.to_csv(
                position_dat_dir + "\\" + f"{skg_name}_ProfileItem-Awards.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
                date_format="%Y/%m/%d",
            )

        else:
            pass

    def profile_item_license(wb, license_df, position_profile_df, position_dat_dir, license_content_item_list):
        ProfileItem_License = wb.sheets[6]
        ProfileItem_License.range((12, "B"), (10000, "Q")).clear()

        spur_id_license_occ = license_df["SPUR ID"].value_counts().sort_index()

        talent_profile_list = []
        license_list = []
        importance_list = []

        row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
        for row in range(12, row_end_talent_profile + 1):
            if wb.sheets[0].range((row, "K")).value == None:
                break

            ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
            ur_id_license_df = license_df[license_df["SPUR ID"] == ur_id]
            position_id = str(re.search("[^_]+$", wb.sheets[0].range((row, "D")).value).group()).zfill(8)
            role_level_column = [
                x for x in position_profile_df.columns if re.search("role level|role$", x.lower().strip())
            ][0]
            role_level = position_profile_df[position_profile_df["Pos ID"] == position_id][role_level_column].values[0]
            jg_column = [x for x in position_profile_df.columns if "jg" in x.lower()][0]
            conso_JG = position_profile_df[position_profile_df["Pos ID"] == position_id][jg_column].values[0]
            # conso_JG = WS_df[WS_df["Pos ID (as per ZHPLA 3/5/2021)"] == position_id]["Conso JG"].iloc[0]
            conso_JG = re.sub(r"Est.|Eqv.", r"", conso_JG).strip()

            if ur_id in spur_id_license_occ.index:
                if all(ur_id_license_df["JG"].notna()) and any([str(conso_JG) in x for x in ur_id_license_df["JG"]]):
                    ur_id_license_df = ur_id_license_df.iloc[[conso_JG in x for x in ur_id_license_df["JG"]]]
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * len(ur_id_license_df)
                    )

                else:
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_license_occ[ur_id]
                    )

            else:
                continue

            if ur_id in license_df["SPUR ID"].values.tolist():
                if all(ur_id_license_df["JG"].notna()) and any([str(conso_JG) in x for x in ur_id_license_df["JG"]]):
                    license_list.extend(
                        ur_id_license_df[ur_id_license_df["SPUR ID"] == ur_id][[license_column]].values.tolist()
                    )
                    importance_list.extend(
                        ur_id_license_df[ur_id_license_df["SPUR ID"] == ur_id][["Importance"]].values.tolist()
                    )

                else:
                    license_list.extend(
                        ur_id_license_df[ur_id_license_df["SPUR ID"] == ur_id][[license_column]].values.tolist()
                    )
                    importance_list.extend(
                        ur_id_license_df[ur_id_license_df["SPUR ID"] == ur_id][["Importance"]].values.tolist()
                    )

            else:
                continue

        if license_list != []:
            # Column H
            ProfileItem_License.range((12, "H")).value = [["={}".format(k)] for k in talent_profile_list]

            # Column E
            ProfileItem_License.range((12, "E")).value = license_list

            # Column J
            # ProfileItem_License.range((12, "J")).value = importance_list

            row_end_license_sheet = (
                ProfileItem_License.range("H" + str(wb.sheets[6].cells.last_cell.row)).end("up").row
            )
            # Column BCDFGIJLMNOP
            for k in "BCDFGIJLMNOP":
                ProfileItem_License.range("{}12:{}{}".format(k, k, row_end_license_sheet)).value = "={}$11".format(k)
            # Column M
            z_list = [
                (
                    ProfileItem_License.range((i, "H"))
                    .formula.replace("=", "")
                    .replace("TalentProfile!K", "TalentProfile!D"),
                    ProfileItem_License.range((i, "D")).get_address(False, False),
                    ProfileItem_License.range((i, "E")).get_address(False, False),
                )
                for i in range(12, row_end_license_sheet + 1)
            ]
            formula_list = [
                [
                    '=CONCAT(UPPER(SUBSTITUTE({}," ","_")),"_",{},"_",UPPER(SUBSTITUTE({}," ","_")),"_PI")'.format(
                        k[0], k[1], k[2]
                    )
                ]
                for k in z_list
            ]
            ProfileItem_License.range((12, "Q")).value = formula_list

            license_content_item_list = license_content_item_list.values.tolist()

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ProfileItem_License.range("B2:Q2").value).T
            ProfileItem_License_df = pd.DataFrame(
                ProfileItem_License.range("B12:Q{}".format(row_end_license_sheet)).value
            )
            ProfileItem_License_df[4] = "2021/11/01"
            ProfileItem_License_df[5] = "4712/12/31"
            # ProfileItem_License_df[8] = ProfileItem_License_df[8].astype(int)
            # ProfileItem_License_df = ProfileItem_License_df.drop_duplicates(subset=[15], keep="first")

            ProfileItem_License.range((12, "B"), (10000, "Q")).clear()
            ProfileItem_License.range((12, "B")).value = ProfileItem_License_df.values.tolist()
            row_end_license_sheet_2 = (
                ProfileItem_License.range("H" + str(ProfileItem_License.cells.last_cell.row)).end("up").row
            )

            ProfileItem_License.range((12, "F"), (row_end_license_sheet_2, "F")).value = "'2021/11/01"
            ProfileItem_License.range((12, "G"), (row_end_license_sheet_2, "G")).value = "'4712/12/31"

            for row in range(12, row_end_license_sheet_2 + 1):
                if not ProfileItem_License.range((row, "E")).value in license_content_item_list:
                    ProfileItem_License.range((row, "E")).color = (255, 0, 0)

            ProfileItem_License_df = pd.concat([header, ProfileItem_License_df])
            ProfileItem_License_df.to_csv(
                position_dat_dir + "\\" + f"{skg_name}_ProfileItem-License.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

        else:
            pass

    # Execute functions
    log.info("[Position profile] Talent Profile")
    talent_profile(wb, position_profile_df, position_dat_dir)

    log.info("[Position profile] Profile Relation")
    profile_relation(wb, position_dat_dir)

    log.info("[Position profile] Model Profile Info")
    model_profile_info(wb, spur_df, position_dat_dir)

    log.info("[Position profile] Profile Attachment")
    profile_attachment(wb, spur_df, position_dat_dir, position_blob_dir)

    # log.info("[Position profile] Profile Item Other Descriptor")
    # profile_item_other_descriptor(wb, spur_df, position_dat_dir)

    log.info("[Position profile] License & Certificate")
    profile_item_license(wb, license_df, position_profile_df, position_dat_dir, license_content_item_list)

    log.info("[Position profile] Degree")
    profile_item_degree(wb, degree_df, position_dat_dir, AreaOfStudy_content_item_list)

    log.info("[Position profile] Honors & Awards")
    profile_item_awards(wb, awards_df, position_profile_df, position_dat_dir, awards_content_item_list)

    log.info("[Position profile] Language")
    profile_item_language(wb, position_dat_dir)

    log.info("[Position profile] Membership")
    profile_item_membership(
        wb,
        membership_df,
        position_profile_df,
        position_dat_dir,
        membership_content_item_list,
    )

    log.info("[Position profile] Experience Required")
    profile_item_exp_required(wb, experience_df, position_profile_df, position_dat_dir)

    log.info("[Position profile] Leadership Competency")
    profile_item_competency_LC(wb, LC_file_path, spur_df, position_profile_df, position_dat_dir, LC_content_item_list)

    log.info("[Position profile] Technical Competency")
    profile_item_competency_TC(
        wb,
        TC_file_path,
        jcp_file_path,
        position_profile_df,
        position_dat_dir,
        TC_content_item_list,
        read_jcp,
    )

    log.info("[Position profile] Profile Item Risk")
    profile_item_risk(wb, spur_df, position_dat_dir)