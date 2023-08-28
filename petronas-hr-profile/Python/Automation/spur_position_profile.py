import pandas as pd
import numpy as np
import re
import xlwings as xw
import os
from config import *

def spur_position_profile(
    position_blob_dir,
    position_template_file_path,
    position_profile_df,
    position_details_file_path,
    position_dat_dir
):
    try:
        run_all = False

        wb = xw.Book(position_template_file_path)

        experience_df = pd.read_excel(position_details_file_path, sheet_name="Experience")
        degree_df = pd.read_excel(position_details_file_path, sheet_name="Degree")
        membership_df = pd.read_excel(position_details_file_path, sheet_name="Membership")
        awards_df = pd.read_excel(position_details_file_path, sheet_name="Awards")
        license_df = pd.read_excel(position_details_file_path, sheet_name="License")
        language_df = pd.read_excel(position_details_file_path, sheet_name="Language")
        leadership_competency_df = pd.read_excel(position_details_file_path, sheet_name="LeadershipCompetency")
        technical_competency_df = pd.read_excel(position_details_file_path, sheet_name="TechnicalCompetency")

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
            talent_profile.range((12, "D")).value = [[x] for x in position_profile_df["PositionProfileCode"].values.tolist()]
            talent_profile.range((12, "E")).value = [[x] for x in position_profile_df["Status"].values.tolist()]
            row_end_talent_profile = talent_profile.range("D" + str(talent_profile.cells.last_cell.row)).end("up").row

            # Column H
            talent_profile.range((12, "H")).value = [
                [x] for x in position_profile_df["Position"].replace("–", "-").values.tolist()
            ]

            # Column I
            talent_profile.range((12, "I"), (row_end_talent_profile, "I")).value = "=H12"

            # Column BCEFGJ
            for col in "BCFGJ":
                talent_profile.range((12, col), (row_end_talent_profile, col)).value = "={}$11".format(col)

            talent_profile.range(
                (12, "K"), (row_end_talent_profile, "K")
            ).value = '=CONCAT(UPPER(SUBSTITUTE(D12," ","_")), "_POSITION_PROFILE")'

            data_range = talent_profile.range("B12:K{}".format(row_end_talent_profile)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            talent_profile_df = pd.DataFrame(data_range)
            header = pd.DataFrame(talent_profile.range("B1:K1").value).T
            talent_profile_df = pd.concat([header, talent_profile_df])
            talent_profile_df.to_csv(
                position_dat_dir + "\\" + f"TalentProfile.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
                date_format="%Y/%m/%d",
            )

        ## Profile Relation ##
        def profile_relation(wb, position_profile_df, position_dat_dir):
            profile_relation = wb.sheets[1]
            profile_relation.range((11, "B"), (10000, "K")).clear()
            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row - 1
            effective_start_date_list = []
            effective_end_date_list = []
            position_code_list = []
            company_name_list = []

            # Column F
            profile_relation.range((11, "F"), (row_end_talent_profile, "F")).value = "=TalentProfile!D12"

            for PositionProfileCode in position_profile_df["PositionProfileCode"].values.tolist():
                effective_start_date = position_profile_df[position_profile_df["PositionProfileCode"] == PositionProfileCode]["EffectiveStartDate"].values.tolist()
                effective_end_date = position_profile_df[position_profile_df["PositionProfileCode"] == PositionProfileCode]["EffectiveEndDate"].values.tolist()
                position_code = position_profile_df[position_profile_df["PositionProfileCode"] == PositionProfileCode]["PositionCode"].values.tolist()
                company_name = position_profile_df[position_profile_df["PositionProfileCode"] == PositionProfileCode]["CompanyName"].values.tolist()
                effective_start_date_list.extend([["'" + item] for item in effective_start_date])
                effective_end_date_list.extend([["'" + item] for item in effective_end_date])
                position_code_list.extend([["'" + item] for item in position_code])
                company_name_list.extend([["'" + item] for item in company_name])

            row_end_profile_relation = (
                profile_relation.range("F" + str(profile_relation.cells.last_cell.row)).end("up").row
            )

            profile_relation.range("D11:D{}".format(row_end_profile_relation)).value = effective_start_date_list
            profile_relation.range("E11:E{}".format(row_end_profile_relation)).value = effective_end_date_list
            profile_relation.range("H11:H{}".format(row_end_profile_relation)).value = position_code_list
            profile_relation.range("I11:I{}".format(row_end_profile_relation)).value = company_name_list

            # Column BCDEGJ
            for col in "BCGJ":
                profile_relation.range((11, col), (row_end_profile_relation, col)).value = "={}$10".format(col)

            profile_relation.range(
                (11, "K"), (row_end_talent_profile, "K")
            ).value = '=CONCAT(UPPER(SUBSTITUTE(F11," ","_")), "_POSITION_PROFILE_RELATION")'

            # Drop competency with minimum proficiency = 0
            data_range = profile_relation.range("B11:K{}".format(row_end_talent_profile)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            profile_relation_df = pd.DataFrame(data_range)
            header = pd.DataFrame(profile_relation.range("B1:K1").value).T

            profile_relation_df[6] = profile_relation_df[6].astype(str).str.zfill(8)

            profile_relation_df = pd.concat([header, profile_relation_df])
            profile_relation_df.to_csv(
                position_dat_dir + "\\" + f"ProfileRelation.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
                date_format="%Y/%m/%d",
            )

        def model_profile_info(wb, position_profile_df, position_dat_dir):
            ModelProfileExtraInfo = wb.sheets[2]
            ModelProfileExtraInfo.range((12, "B"), (100000, "I")).clear()
            row_end_talent_profile = wb.sheets[0].range("H" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            # Column D
            ModelProfileExtraInfo.range((12, "D"), (row_end_talent_profile, "D")).value = "=TalentProfile!D12"
            row_end_ModelProfileExtraInfo = (
                ModelProfileExtraInfo.range("D" + str(ModelProfileExtraInfo.cells.last_cell.row)).end("up").row
            )

            # Column EFG
            position_id_list = [
                x if x in position_profile_df["PositionProfileCode"].values.tolist() else ""
                for x in ModelProfileExtraInfo.range((12, "D")).expand("down").options(ndim=1).value
            ]
            position_id_txt_list = [
                [
                    x + "_DESCRIPTION.txt",
                    x + "_QUALIFICATION.txt",
                    x + "_RESPONSIBILITY.txt",
                ]
                if x in position_profile_df["PositionProfileCode"].values.tolist()
                else ["", "", ""]
                for x in position_id_list
            ]
            ModelProfileExtraInfo.range((12, "E")).value = position_id_txt_list

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

            data_range = ModelProfileExtraInfo.range("B12:I{}".format(row_end_talent_profile)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ModelProfileExtraInfo_df = pd.DataFrame(data_range)
            header = pd.DataFrame(ModelProfileExtraInfo.range("B1:I1").value).T
            ModelProfileExtraInfo_df = pd.concat([header, ModelProfileExtraInfo_df])
            ModelProfileExtraInfo_df.to_csv(
                position_dat_dir + "\\" + f"ModelProfileExtraInfo.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
                date_format="%Y/%m/%d",
            )

        def profile_attachment(wb, position_profile_df, position_dat_dir, position_blob_dir):
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

                position_profile_code = wb.sheets[0].range((row, "D")).value
                position_name = wb.sheets[0].range((row, "H")).get_address(False, False, True)
                position_profile = position_profile_df[position_profile_df['PositionProfileCode'] == position_profile_code].iloc[0]
                spur_file_name = position_profile_code + '_' + position_profile['SPURProfileCode'].split('_')[0] + '.pdf'
                pd_file_name = position_profile_code + '_' + position_profile_code.split('_')[2] + '.pdf'

                if any(item.startswith(position_profile_code) for item in blob_files_name):
                    profile_code_list.extend(["=" + wb.sheets[0].range((row, "D")).get_address(False, False, True)] * 2)
                    position_name_list.extend(["=" + position_name] * 2)
                    final_blob_files_list.extend([spur_file_name, pd_file_name])

                else:
                    profile_code_list.extend(["=" + wb.sheets[0].range((row, "D")).get_address(False, False, True)])
                    position_name_list.extend(["=" + position_name])
                    final_blob_files_list.extend([pd_file_name])

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

            data_range = ProfileAttachment.range("B11:K{}".format(row_end_ProfileAttachment)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ProfileAttachment_df = pd.DataFrame(data_range)
            header = pd.DataFrame(ProfileAttachment.range("B1:K1").value).T
            ProfileAttachment_df = pd.concat([header, ProfileAttachment_df])
            ProfileAttachment_df.to_csv(
                position_dat_dir + "\\" + f"ProfileAttachment.dat",
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
                for x in ProfileItem_OtherDescriptor.range("H9").expand("down").options(ndim=1).value
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
                position_dat_dir + "\\" + f"ProfileItem-OtherDescriptor.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
                date_format="%Y/%m/%d",
            )

        def profile_item_risk(wb, position_profile_df, position_dat_dir):
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

            # Column J
            ProfileItem_Risk.range((12, "J"), (row_end_ProfileItem_Risk, "J")).value = [
                [
                    position_profile_df[position_profile_df["PositionProfileCode"] == x.replace('_POSITION_PROFILE', '')]["Challenge"].values[0].replace("\n", "")
                    if x.replace('_POSITION_PROFILE', '') in position_profile_df["PositionProfileCode"].values.tolist()
                    else np.nan
                ]
                for x in ProfileItem_Risk.range("G12").expand("down").options(ndim=1).value
            ]

            # Column L
            ProfileItem_Risk.range(
                (12, "L"), (row_end_ProfileItem_Risk, "L")
            ).value = '=CONCATENATE(UPPER(TalentProfile!D12),"_",UPPER(SUBSTITUTE(H12," ","_")),"_PI")'

            data_range = ProfileItem_Risk.range("B12:L{}".format(row_end_ProfileItem_Risk)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ProfileItem_Risk_df = pd.DataFrame(data_range)
            header = pd.DataFrame(ProfileItem_Risk.range("B2:L2").value).T
            ProfileItem_OtherDescriptor_df = pd.concat([header, ProfileItem_Risk_df])
            ProfileItem_OtherDescriptor_df.to_csv(
                position_dat_dir + "\\" + f"ProfileItem-Risk.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
                date_format="%Y/%m/%d",
            )

        def profile_item_exp_required(wb, experience_df, position_dat_dir):
            ProfileItem_ExperienceRequired = wb.sheets[16]
            ProfileItem_ExperienceRequired.range((12, "B"), (10000, "Q")).clear()
            profile_code_occ = experience_df["PositionProfileCode"].value_counts().sort_index()

            id_list = []
            min_exp_list = []
            max_exp_list = []
            industry_list = []
            domain_list = []
            skill_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(12, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                profile_code = wb.sheets[0].range((row, "D")).value

                if profile_code in profile_code_occ.index:
                    id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(False, False, True)] * profile_code_occ[profile_code]
                    )

                else:
                    if run_all == True:
                        id_list.extend([wb.sheets[0].range((row, "K")).get_address(False, False, True)])
                    else:
                        continue

                if profile_code in experience_df["PositionProfileCode"].values.tolist():
                    min_exp_list.extend(
                        experience_df[experience_df["PositionProfileCode"] == profile_code][
                            ['MimimumExperienceRequired']
                        ].values.tolist()
                    )
                    max_exp_list.extend(
                        experience_df[experience_df["PositionProfileCode"] == profile_code][
                            ['MaximumExperienceRequired']
                        ].values.tolist()
                    )
                    
                    industry_list.extend(

                        experience_df[experience_df["PositionProfileCode"] == profile_code][['Industry']].values.tolist()
                    )
                    domain_list.extend(
                        experience_df[experience_df["PositionProfileCode"] == profile_code][
                            ['Domain']
                        ].values.tolist()
                    )
                    skill_list.extend(
                        experience_df[experience_df["PositionProfileCode"] == profile_code][['Skill']].values.tolist()
                    )

                else:
                    if run_all == True:
                        min_exp_list.extend([[""]])
                        max_exp_list.extend([[""]])
                        industry_list.extend([[""]])
                        domain_list.extend([[""]])
                        skill_list.extend([[""]])
                    else:
                        continue
            
            if id_list != []:
                # Column H
                ProfileItem_ExperienceRequired.range((12, "H")).value = [["={}".format(k)] for k in id_list]
                row_end_ProfileItem_ExperienceRequired = (
                    ProfileItem_ExperienceRequired.range("H" + str(ProfileItem_ExperienceRequired.cells.last_cell.row))
                    .end("up")
                    .row
                )

                # Column J
                ProfileItem_ExperienceRequired.range((12, "J")).value = skill_list

                # Column L
                ProfileItem_ExperienceRequired.range((12, "L")).value = min_exp_list

                # Column M
                ProfileItem_ExperienceRequired.range((12, "M")).value = max_exp_list

                # Column N
                ProfileItem_ExperienceRequired.range((12, "N")).value = domain_list

                # Column O
                ProfileItem_ExperienceRequired.range((12, "O")).value = industry_list

                # Column BCDFGIKP
                for k in "BCDFGIKP":
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
                        '=CONCAT(UPPER(SUBSTITUTE({}," ","_")),"_",{},"_",UPPER(SUBSTITUTE({}, " ","_")),"_" ,UPPER(SUBSTITUTE({}, " ", "_")),"_PI")'.format(
                            k[0], k[1], k[2], k[3]
                        )
                    ]
                    for k in z_list
                ]
                ProfileItem_ExperienceRequired.range((12, "Q")).value = formula_list

                header = pd.DataFrame(ProfileItem_ExperienceRequired.range("B2:Q2").value).T
                data_range = ProfileItem_ExperienceRequired.range("B12:Q{}".format(row_end_ProfileItem_ExperienceRequired)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_ExperienceRequired_df = pd.DataFrame(data_range)
                ProfileItem_ExperienceRequired_df = pd.concat([header, ProfileItem_ExperienceRequired_df])
                ProfileItem_ExperienceRequired_df.to_csv(
                    position_dat_dir + "\\" + f"ProfileItem-ExperienceRequired.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                    line_terminator="\n",
                )
            else:
                pass

        def profile_item_competency_LC(
            wb, leadership_competency_df, position_dat_dir
        ):
            ProfileItem_Competency_LC = wb.sheets[20]
            ProfileItem_Competency_LC.range((13, "B"), (100000, "Q")).clear()
            profile_code_leadership_competency_occ = leadership_competency_df["PositionProfileCode"].value_counts().sort_index()
            lc_id_list = []
            lc_list = []
            min_list = []
            max_list = []
            
            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(12, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                profile_code = wb.sheets[0].range((row, "D")).value
                #     print(ur_id)

                if profile_code in profile_code_leadership_competency_occ.index:
                    lc_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * profile_code_leadership_competency_occ[profile_code]
                    )

                else:
                    if run_all == True:
                        lc_id_list.extend([wb.sheets[0].range((row, "K")).get_address(True, False, True)])
                    else:
                        continue

                if profile_code in leadership_competency_df["PositionProfileCode"].values.tolist():
                    lc_list.extend(leadership_competency_df[leadership_competency_df["PositionProfileCode"] == profile_code][['LeadershipCompetencyName']].values.tolist())
                    min_list.extend(leadership_competency_df[leadership_competency_df["PositionProfileCode"] == profile_code][['MinimumProficiency']].values.tolist())
                    max_list.extend(leadership_competency_df[leadership_competency_df["PositionProfileCode"] == profile_code][['MaximumProficiency']].values.tolist())
                else:
                    continue

            if lc_id_list != []:

                ProfileItem_Competency_LC.range((13, "H")).value = [["={}".format(k)] for k in lc_id_list]

                ProfileItem_Competency_LC.range((13, "E")).value = lc_list
                ProfileItem_Competency_LC.range((13, "K")).value = min_list
                ProfileItem_Competency_LC.range((13, "M")).value = max_list

                last_row_lc = (
                    ProfileItem_Competency_LC.range("H" + str(ProfileItem_Competency_LC.cells.last_cell.row)).end("up").row
                )

                # Column BCDFGIJLN
                for col in "BCDFGIJLN":
                    ProfileItem_Competency_LC.range((13, col), (last_row_lc, col)).value = "={}$11".format(col)

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
                header = pd.DataFrame(ProfileItem_Competency_LC.range("B2:O2").value).T
                data_range = ProfileItem_Competency_LC.range("B13:O{}".format(last_row_lc)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_Competency_LC_df = pd.DataFrame(data_range)
                ProfileItem_Competency_LC_df[[9, 11]] = ProfileItem_Competency_LC_df[[9, 11]].apply(
                    pd.to_numeric, downcast="signed"
                )

                ProfileItem_Competency_LC_df = pd.concat([header, ProfileItem_Competency_LC_df])
                ProfileItem_Competency_LC_df.to_csv(
                    position_dat_dir + "\\" + f"ProfileItem-Competency_LC.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                )
            else:
                pass

        def profile_item_competency_TC(wb,technical_competency_df, position_dat_dir):
            ProfileItem_Competency_TC = wb.sheets[21]
            ProfileItem_Competency_TC.range((13, "B"), (10000, "P")).clear()

            profile_code_technical_competency_occ = technical_competency_df["PositionProfileCode"].value_counts().sort_index()
            tc_id_list = []
            tc_list = []
            min_list = []
            max_list = []
            importance_list = []
            
            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(12, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                profile_code = wb.sheets[0].range((row, "D")).value
                #     print(ur_id)

                if profile_code in profile_code_technical_competency_occ.index:
                    tc_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * profile_code_technical_competency_occ[profile_code]
                    )

                else:
                    if run_all == True:
                        tc_id_list.extend([wb.sheets[0].range((row, "K")).get_address(True, False, True)])
                    else:
                        continue

                if profile_code in technical_competency_df["PositionProfileCode"].values.tolist():
                    tc_list.extend(technical_competency_df[technical_competency_df["PositionProfileCode"] == profile_code][['TechnicalCompetencyName']].values.tolist())
                    min_list.extend(technical_competency_df[technical_competency_df["PositionProfileCode"] == profile_code][['MinimumProficiency']].values.tolist())
                    max_list.extend(technical_competency_df[technical_competency_df["PositionProfileCode"] == profile_code][['MaximumProficiency']].values.tolist())
                    importance_list.extend(technical_competency_df[technical_competency_df["PositionProfileCode"] == profile_code][['Importance']].values.tolist())
                else:
                    continue
            
            if tc_id_list != []:
                # Column H
                ProfileItem_Competency_TC.range((13, "H")).value = [["={}".format(k)] for k in tc_id_list]

                # Column E
                ProfileItem_Competency_TC.range((13, "E")).value = tc_list

                # Column E
                ProfileItem_Competency_TC.range((13, "L")).value = min_list

                # Column N
                ProfileItem_Competency_TC.range((13, "N")).value = max_list

                # column J
                ProfileItem_Competency_TC.range((13, "J")).value = importance_list

                last_row_tc = 13 + len(tc_id_list) - 1
                # Column BCDFGIKMO
                for col in "BCDFGIKMO":
                    ProfileItem_Competency_TC.range((13, col), (last_row_tc, col)).value = "={}$11".format(col)

                # Column P
                z_list = [
                    (
                        ProfileItem_Competency_TC.range((i, "H"))
                        .formula.replace("=", "")
                        .replace("TalentProfile!K", "TalentProfile!D"),
                        ProfileItem_Competency_TC.range((i, "I")).get_address(False, False),
                        ProfileItem_Competency_TC.range((i, "E")).get_address(False, False),
                    )
                    for i in range(13, last_row_tc + 1)
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

                header = pd.DataFrame(ProfileItem_Competency_TC.range("B2:P2").value).T
                data_range = ProfileItem_Competency_TC.range("B13:P{}".format(last_row_tc)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_Competency_TC_df = pd.DataFrame(data_range).dropna(subset=[10])
                ProfileItem_Competency_TC_df[8] = ProfileItem_Competency_TC_df[8].fillna(-1).astype(int).replace(-1, "")
                ProfileItem_Competency_TC_df[10] = ProfileItem_Competency_TC_df[10].fillna(-1).astype(int).replace(-1, "")
                ProfileItem_Competency_TC_df[12] = ProfileItem_Competency_TC_df[12].fillna(-1).astype(int).replace(-1, "")
                ProfileItem_Competency_TC_df = pd.concat([header, ProfileItem_Competency_TC_df])
                
                ProfileItem_Competency_TC_df.to_csv(
                    position_dat_dir + "\\" + f"ProfileItem-Competency_TC.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                )
            else:
                pass

        def profile_item_degree(wb, degree_df, position_dat_dir):
            ProfileItem_Degree = wb.sheets[8]
            ProfileItem_Degree.range((12, "B"), (10000, "Q")).clear()

            profile_code_degree_occ = degree_df["PositionProfileCode"].value_counts().sort_index()

            degree_id_list = []
            edu_level_list = []
            degree_name_list = []
            country_code_list = []
            required_list = []
            major_list = []
            school_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(12, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                profile_code = wb.sheets[0].range((row, "D")).value

                if profile_code in profile_code_degree_occ.index:
                    degree_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * profile_code_degree_occ[profile_code]
                    )

                else:
                    continue

                if profile_code in degree_df["PositionProfileCode"].values.tolist():
                    edu_level_list.extend(degree_df[degree_df["PositionProfileCode"] == profile_code][['DegreeName']].values.tolist())
                    degree_name_list.extend(
                        degree_df[degree_df["PositionProfileCode"] == profile_code][['StudyAreaName']].values.tolist()
                    )
                    country_code_list.extend(
                        degree_df[degree_df["PositionProfileCode"] == profile_code][['CountryCode']].values.tolist()
                    )
                    required_list.extend(
                        degree_df[degree_df["PositionProfileCode"] == profile_code][['Required']].values.tolist()
                    )
                    major_list.extend(
                        degree_df[degree_df["PositionProfileCode"] == profile_code][['Major']].values.tolist()
                    )
                    school_list.extend(
                        degree_df[degree_df["PositionProfileCode"] == profile_code][['School']].values.tolist()
                    )
                else:
                    if run_all == True:
                        edu_level_list.extend([[""]])
                        degree_name_list.extend([[""]])
                        country_code_list.extend([[""]])
                        required_list.extend([[""]])
                        major_list.extend([[""]])
                        school_list.extend([[""]])
                    else:
                        continue

            if degree_id_list != []:
                # Column H
                ProfileItem_Degree.range((12, "H")).value = [["={}".format(k)] for k in degree_id_list]

                # Column J
                ProfileItem_Degree.range((12, "J")).value = required_list

                # Column K
                ProfileItem_Degree.range((12, "K")).value = major_list

                # Column L
                ProfileItem_Degree.range((12, "L")).value = school_list

                # Column E
                ProfileItem_Degree.range((12, "E")).value = edu_level_list

                # Column O
                ProfileItem_Degree.range((12, "O")).value = degree_name_list

                # Column P
                ProfileItem_Degree.range((12, "P")).value = degree_name_list

                # Column N
                ProfileItem_Degree.range((12, "N")).value = country_code_list

                # Column O
                ProfileItem_Degree.range((12, "O")).value = country_code_list

                row_end_degree_sheet = (
                    ProfileItem_Degree.range("H" + str(ProfileItem_Degree.cells.last_cell.row)).end("up").row
                )
                # Column BCDFGILMNP
                for k in "BCDFGIMQ":
                    ProfileItem_Degree.range("{}12:{}{}".format(k, k, row_end_degree_sheet)).value = "={}$11".format(k)

                # Column F
                ProfileItem_Degree.range((12, "F"), (row_end_degree_sheet, "F")).value = "'2021/11/01"

                # Column G
                ProfileItem_Degree.range((12, "G"), (row_end_degree_sheet, "G")).value = "'4712/12/31"

                # Column Q
                z_list = [
                    (
                        ProfileItem_Degree.range((i, "H"))
                        .formula.replace("=", "")
                        .replace("TalentProfile!K", "TalentProfile!D"),
                        ProfileItem_Degree.range((i, "I")).get_address(False, False),
                        ProfileItem_Degree.range((i, "E")).get_address(False, False),
                        ProfileItem_Degree.range((i, "O")).get_address(False, False),
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

                # Drop competency with minimum proficiency = 0
                header = pd.DataFrame(ProfileItem_Degree.range("B2:R2").value).T
                data_range = ProfileItem_Degree.range("B12:R{}".format(row_end_degree_sheet)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_Degree_df = pd.DataFrame(data_range)
                ProfileItem_Degree_df = pd.concat([header, ProfileItem_Degree_df])
                ProfileItem_Degree_df.to_csv(
                    position_dat_dir + "\\" + f"ProfileItem-Degree.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                    float_format="%.f",
                )
            else:
                pass

        def profile_item_language(wb, language_df, position_dat_dir):
            ProfileItem_Language = wb.sheets[12]
            ProfileItem_Language.range((12, "B"), (10000, "R")).clear()

            profile_code_language_occ = language_df["PositionProfileCode"].value_counts().sort_index()
            language_id_list = []
            language_list = []
            reading_proficiency_list = []
            writing_proficiency_list = []
            speaking_proficiency_list = []
            required_list = []
            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            for row in range(12, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                profile_code = wb.sheets[0].range((row, "D")).value

                if profile_code in profile_code_language_occ.index:
                    language_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * profile_code_language_occ[profile_code]
                    )

                else:
                    continue

                if profile_code in language_df["PositionProfileCode"].values.tolist():
                    language_list.extend(language_df[language_df["PositionProfileCode"] == profile_code][['LanguageName']].values.tolist())
                    reading_proficiency_list.extend(language_df[language_df["PositionProfileCode"] == profile_code][['ReadingProficiency']].values.tolist())
                    writing_proficiency_list.extend(language_df[language_df["PositionProfileCode"] == profile_code][['WritingProficiency']].values.tolist())
                    speaking_proficiency_list.extend(language_df[language_df["PositionProfileCode"] == profile_code][['SpeakingProficiency']].values.tolist())
                    required_list.extend(language_df[language_df["PositionProfileCode"] == profile_code][['Required']].values.tolist())

                else:
                    continue

            if language_id_list != []:
                # Column H
                ProfileItem_Language.range((12, "H")).value = [["={}".format(k)] for k in language_id_list]

                ProfileItem_Language.range((12, "E")).value = language_list
                ProfileItem_Language.range((12, "L")).value = reading_proficiency_list
                ProfileItem_Language.range((12, "N")).value = writing_proficiency_list
                ProfileItem_Language.range((12, "P")).value = speaking_proficiency_list
                ProfileItem_Language.range((12, "K")).value = required_list

                row_end_language_sheet = ProfileItem_Language.range("H" + str(wb.sheets[11].cells.last_cell.row)).end("up").row
                # Column BCDEFGIJKLMNOPQ
                for k in "BCDFGIJMOQ":
                    ProfileItem_Language.range("{}12:{}{}".format(k, k, row_end_language_sheet)).value = "={}$11".format(k)

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
                data_range = ProfileItem_Language.range("B12:R{}".format(row_end_language_sheet)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_Language_df = pd.DataFrame(data_range)
                ProfileItem_Language_df = pd.concat([header, ProfileItem_Language_df])
                ProfileItem_Language_df.to_csv(
                    position_dat_dir + "\\" + f"ProfileItem-Language.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                )
            else:
                pass

        def profile_item_membership(
            wb,
            membership_df,
            position_dat_dir
        ):
            ProfileItem_Membership = wb.sheets[14]
            ProfileItem_Membership.range((12, "B"), (10000, "N")).clear()

            profile_code_membership_occ = membership_df["PositionProfileCode"].value_counts().sort_index()

            talent_profile_list = []
            membership_list = []
            required_list = []
            title_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(12, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                profile_code = wb.sheets[0].range((row, "D")).value

                if profile_code in profile_code_membership_occ.index:
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * profile_code_membership_occ[profile_code]
                    )

                else:
                    continue

                if profile_code in membership_df["PositionProfileCode"].values.tolist():
                    membership_list.extend(
                        membership_df[membership_df["PositionProfileCode"] == profile_code][['MembershipName']].values.tolist()
                    )
                    required_list.extend(
                        membership_df[membership_df["PositionProfileCode"] == profile_code][['Required']].values.tolist()
                    )
                    title_list.extend(
                        membership_df[membership_df["PositionProfileCode"] == profile_code][['Title']].values.tolist()
                    )
                else:
                    continue

            if membership_list != []:
                # Column H
                ProfileItem_Membership.range((12, "H")).value = [["={}".format(k)] for k in talent_profile_list]

                # Column E
                ProfileItem_Membership.range((12, "E")).value = membership_list

                # Column J
                ProfileItem_Membership.range((12, "J")).value = required_list

                # Column K
                ProfileItem_Membership.range((12, "K")).value = title_list

                row_end_membership_sheet = (
                    ProfileItem_Membership.range("H" + str(wb.sheets[13].cells.last_cell.row)).end("up").row
                )
                # Column BCDFGIM
                for k in "BCDFGILM":
                    ProfileItem_Membership.range(
                        "{}12:{}{}".format(k, k, row_end_membership_sheet)
                    ).value = "={}$11".format(k)
                # Column M
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

                data_range = ProfileItem_Membership.range("B12:N{}".format(row_end_membership_sheet)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_Membership_df = pd.DataFrame(data_range)
                header = pd.DataFrame(ProfileItem_Membership.range("B2:N2").value).T
                ProfileItem_Membership_df = pd.concat([header, ProfileItem_Membership_df])
                ProfileItem_Membership_df.to_csv(
                    position_dat_dir + "\\" + f"ProfileItem-Membership.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                    date_format="%Y/%m/%d",
                )

            else:
                pass

        def profile_item_awards(wb, awards_df, position_dat_dir):
            ProfileItem_Awards = wb.sheets[10]
            ProfileItem_Awards.range((12, "B"), (10000, "N")).clear()

            profile_code_awards_occ = awards_df["PositionProfileCode"].value_counts().sort_index()

            talent_profile_list = []
            awards_list = []
            required_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(12, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                profile_code = wb.sheets[0].range((row, "D")).value

                if profile_code in profile_code_awards_occ.index:
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * profile_code_awards_occ[profile_code]
                    )

                else:
                    continue

                if profile_code in awards_df["PositionProfileCode"].values.tolist():
                    awards_list.extend(awards_df[awards_df["PositionProfileCode"] == profile_code][["AwardName"]].values.tolist())
                    required_list.extend(awards_df[awards_df["PositionProfileCode"] == profile_code][["Required"]].values.tolist())

                else:
                    continue

            if awards_list != []:
                # Column H
                ProfileItem_Awards.range((12, "H")).value = [["={}".format(k)] for k in talent_profile_list]

                # Column E
                ProfileItem_Awards.range((12, "E")).value = awards_list

                # Column J
                ProfileItem_Awards.range((12, "J")).value = required_list

                row_end_awards_sheet = ProfileItem_Awards.range("H" + str(wb.sheets[9].cells.last_cell.row)).end("up").row
                # Column BCDFGIJKM
                for k in "BCDFGIKL":
                    ProfileItem_Awards.range("{}12:{}{}".format(k, k, row_end_awards_sheet)).value = "={}$11".format(k)
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
                ProfileItem_Awards.range((12, "M")).value = formula_list

                header = pd.DataFrame(ProfileItem_Awards.range("B2:N2").value).T
                data_range = ProfileItem_Awards.range("B12:N{}".format(row_end_awards_sheet)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_Awards_df = pd.DataFrame(data_range)
                ProfileItem_Awards_df = pd.concat([header, ProfileItem_Awards_df])
                ProfileItem_Awards_df.to_csv(
                    position_dat_dir + "\\" + f"ProfileItem-Awards.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                )

            else:
                pass

        def profile_item_license(wb, license_df, position_dat_dir):
            ProfileItem_License = wb.sheets[6]
            ProfileItem_License.range((12, "B"), (10000, "Q")).clear()

            profile_code_license_occ = license_df["PositionProfileCode"].value_counts().sort_index()

            talent_profile_list = []
            license_list = []
            required_list = []
            country_code_list = []
            state_name_list = []
            title_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(12, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "D")).value == None:
                    break

                profile_code = wb.sheets[0].range((row, "D")).value

                if profile_code in profile_code_license_occ.index:
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * profile_code_license_occ[profile_code]
                    )

                else:
                    continue

                if profile_code in license_df["PositionProfileCode"].values.tolist():
                    license_list.extend(license_df[license_df["PositionProfileCode"] == profile_code][['LicenseName']].values.tolist())
                    required_list.extend(license_df[license_df["PositionProfileCode"] == profile_code][['Required']].values.tolist())
                    country_code_list.extend(license_df[license_df["PositionProfileCode"] == profile_code][['CountryCode']].values.tolist())
                    state_name_list.extend(license_df[license_df["PositionProfileCode"] == profile_code][['StateName']].values.tolist())
                    title_list.extend(license_df[license_df["PositionProfileCode"] == profile_code][['Title']].values.tolist())
                else:
                    continue

            if license_list != []:
                # Column H
                ProfileItem_License.range((12, "H")).value = [["={}".format(k)] for k in talent_profile_list]

                # Column E
                ProfileItem_License.range((12, "E")).value = license_list

                # Column K
                ProfileItem_License.range((12, "K")).value = title_list

                # Column J
                ProfileItem_License.range((12, "J")).value = required_list

                # Column L
                ProfileItem_License.range((12, "L")).value = country_code_list

                # Column M
                ProfileItem_License.range((12, "M")).value = country_code_list

                # Column N
                ProfileItem_License.range((12, "N")).value = country_code_list

                # Column O
                ProfileItem_License.range((12, "O")).value = state_name_list

                row_end_license_sheet = (
                    ProfileItem_License.range("H" + str(wb.sheets[5].cells.last_cell.row)).end("up").row
                )
                # Column BCDFGIJLMNOP
                for k in "BCDFGIP":
                    ProfileItem_License.range("{}12:{}{}".format(k, k, row_end_license_sheet)).value = "={}$11".format(k)
                # Column Q
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

                header = pd.DataFrame(ProfileItem_License.range("B2:Q2").value).T
                data_range = ProfileItem_License.range("B12:Q{}".format(row_end_license_sheet)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_License_df = pd.DataFrame(data_range)
                ProfileItem_License_df = pd.concat([header, ProfileItem_License_df])
                ProfileItem_License_df.to_csv(
                    position_dat_dir + "\\" + f"ProfileItem-License.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                )

            else:
                pass

        # Execute functions
        talent_profile(wb, position_profile_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] Talent Profile.")

        profile_relation(wb, position_profile_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] Profile Relation.")

        model_profile_info(wb, position_profile_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] Model Profile Info.")

        profile_attachment(wb, position_profile_df, position_dat_dir, position_blob_dir)
        LOGGER.info("Completed [Position profile] Profile Attachment.")

        # log.info("[Position profile] Profile Item Other Descriptor")
        # profile_item_other_descriptor(wb, spur_df, position_dat_dir)

        profile_item_license(wb, license_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] License & Certificate.")

        profile_item_degree(wb, degree_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] Degree.")

        profile_item_awards(wb, awards_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] Honors & Awards.")

        profile_item_language(wb, language_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] Language.")

        profile_item_membership(
            wb,
            membership_df,
            position_dat_dir
        )
        LOGGER.info("Completed [Position profile] Membership.")

        profile_item_exp_required(wb, experience_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] Experience Required.")

        profile_item_competency_LC(wb, leadership_competency_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] Leadership Competency.")

        profile_item_competency_TC(
            wb,
            technical_competency_df,
            position_dat_dir
        )
        LOGGER.info("Completed [Position profile] Technical Competency.")

        profile_item_risk(wb, position_profile_df, position_dat_dir)
        LOGGER.info("Completed [Position profile] Profile Item Risk.")
    except Exception as e:
        raise ValueError(e)
    finally:
        if 'wb' in locals():
            wb.save()
            wb.app.quit()