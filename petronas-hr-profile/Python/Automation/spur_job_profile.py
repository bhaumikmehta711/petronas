import pandas as pd
import numpy as np
import re
import xlwings as xw
import logging
import os


def spur_job_profile(
    job_template_file_path,
    spur_df,
    spur_details_file_path,
    job_dat_dir,
    log_dir,
):
    try:

        run_all = False

        log = logging.getLogger(__name__)

        wb = xw.Book(job_template_file_path)

        def profile_code_map(string):
            """
            This function map role level of UR into designation
            """
            if not isinstance(string, str):
                return np.nan

            if "chief" in string.lower():
                return "SGM"

            elif "general manager" in string.lower() or "custodian" in string.lower() or "head" in string.lower():
                return "GM"

            elif "senior manager" in string.lower() or "principal" in string.lower():
                return "SM"

            elif "manager" in string.lower() or "staff" in string.lower():
                return "MANAGER"

            elif "executive" in string.lower():
                return "EXECUTIVE"

            else:
                return ""

        def role_level_jg_map(grade):
            """
            docstring
            """
            # if not isinstance(grade, str):
            #     return np.nan

            grade = str(grade)

            if re.search("H1|H2", grade):
                return "CHIEF"

            if re.search("C1|C2", grade):
                return "GM"

            if re.search("M1|M2", grade):
                return "SM"

            if re.search("D2|D3", grade):
                return "MANAGER"

            if re.search("A1|A2|A3|D1", grade):
                return "EXECUTIVE"

            if re.search("E3", grade):
                return "STAFF"

            if re.search("E4", grade):
                return "PRINCIPAL"

            if re.search("E5", grade):
                return "CUSTODIAN"

            else:
                pass


        spur_df["sort"] = spur_df["UR_CODE"].str.extract("([\w]+$)", expand=False).str.replace("a", "").astype(int)
        spur_df.sort_values("sort", inplace=True, ascending=True)
        # remove duplicates
        spur_df = spur_df.drop_duplicates()
        spur_df = spur_df.drop("sort", axis=1)

        experience_df = pd.read_excel(spur_details_file_path, sheet_name="Experience").replace("\n", "<br>", regex=True)
        experience_df_obj = experience_df.select_dtypes(["object"])
        experience_df[experience_df_obj.columns] = experience_df_obj.apply(lambda x: x.str.strip())
        ur_name_map = dict(zip(spur_df["UR_CODE"], spur_df["UR_NAME"]))
        experience_df["UR_NAME"] = experience_df["SPUR ID"].map(ur_name_map)
        # experience_df["role_level"] = experience_df["JG"].apply(lambda x: role_level_jg_map(x))
        # experience_df["role_level"] = experience_df["role_level"].fillna(
        #     experience_df["UR_NAME"].apply(lambda x: profile_code_map(x))
        # )
        experience_df["profile_code"] = experience_df["SPUR ID"].str.replace(" ", "_") + "_" + experience_df["Role Level"]

        industry_column = [x for x in experience_df.columns if re.search("Industry|Field", x, flags=re.I)][0]
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
        exp_jg_column = [x for x in experience_df.columns if re.search("JG", x, flags=re.I)][0]
        # data = []
        # for spur_id in experience_df['SPUR ID'].unique():
        #     temp_df = experience_df[experience_df['SPUR ID'] == spur_id]
        #     for field in temp_df[industry_column].unique():
        #         min_years = temp_df[temp_df[industry_column] == field][min_years_column].min()
        #         max_years = temp_df[temp_df[industry_column] == field][max_years_column].max()
        #         importance = temp_df[temp_df[industry_column] == field]['Importance'].iloc[0]
        #         if domain_exist == True:
        #             domain = temp_df[temp_df[industry_column] == field][domain_column].iloc[0]
        #             data.append([spur_id, min_years, max_years, field, domain, importance])
        #         else:
        #             data.append([spur_id, min_years, max_years, field, importance])
        # if domain_exist == True:
        #     experience_df = pd.DataFrame(data, columns=['SPUR ID', min_years_column, max_years_column, industry_column, domain_column, 'Importance'])
        # else:
        #     experience_df = pd.DataFrame(data, columns=['SPUR ID', min_years_column, max_years_column, industry_column, 'Importance'])

        degree_df = pd.read_excel(spur_details_file_path, sheet_name="Degree")
        degree_df_obj = degree_df.select_dtypes(["object"])
        degree_df[degree_df_obj.columns] = degree_df_obj.apply(lambda x: x.str.strip())
        degree_column = [x for x in degree_df.columns if re.search("Degree|ContentItem", str(x), flags=re.I)][0]
        area_of_study_column = [
            x for x in degree_df.columns if re.search("Area of study|AreaOfStudy", str(x), flags=re.I)
        ][0]
        degree_df = degree_df.drop_duplicates(subset=["SPUR ID", degree_column, area_of_study_column], keep="first")

        membership_df = pd.read_excel(spur_details_file_path, sheet_name="Membership")
        membership_df_obj = membership_df.select_dtypes(["object"])
        membership_df[membership_df_obj.columns] = membership_df_obj.apply(lambda x: x.str.strip())
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
        awards_df_obj = awards_df.select_dtypes(["object"])
        awards_df[awards_df_obj.columns] = awards_df_obj.apply(lambda x: x.str.strip())

        license_df = pd.read_excel(spur_details_file_path, sheet_name="License")
        license_df_obj = license_df.select_dtypes(["object"])
        license_df[license_df_obj.columns] = license_df_obj.apply(lambda x: x.str.strip())
        license_column = [x for x in license_df.columns if re.search("License", x, flags=re.I)][0]

        language_df = pd.read_excel(spur_details_file_path, sheet_name="Language")
        language_df_obj = language_df.select_dtypes(["object"])
        language_df[language_df_obj.columns] = language_df_obj.apply(lambda x: x.str.strip())
        language_column = [x for x in language_df.columns if re.search("Language", x, flags=re.I)][0]

        leadership_competency_df = pd.read_excel(spur_details_file_path, sheet_name="LeadershipCompetency")
        leadership_competency_df_obj = leadership_competency_df.select_dtypes(["object"])
        leadership_competency_df[leadership_competency_df_obj.columns] = leadership_competency_df_obj.apply(lambda x: x.str.strip())
        leadership_competency_column = [x for x in leadership_competency_df.columns if re.search("Competency", x, flags=re.I)][0]

        technical_competency_df = pd.read_excel(spur_details_file_path, sheet_name="TechnicalCompetency")
        technical_competency_df_obj = technical_competency_df.select_dtypes(["object"])
        technical_competency_df[technical_competency_df_obj.columns] = technical_competency_df_obj.apply(lambda x: x.str.strip())
        technical_competency_column = [x for x in technical_competency_df.columns if re.search("Competency", x, flags=re.I)][0]

        # content_item_dict = pd.read_excel(content_item_file_path, sheet_name=None)
        # LC_content_item_list = (
        #     content_item_dict["ContentItem-Competency Edge"]
        #     .loc[9:, "Content Type Name"]
        #     .str.strip()
        #     .apply(lambda x: " ".join(str(x).split()))
        # )
        # # competency_tec_sheet = [
        # #     x for x in content_item_dict.keys() if re.search("ContentItem-Competency$", x, flags=re.I)
        # # ][0]
        # # TC_content_item_list = (
        # #     content_item_dict[competency_tec_sheet]
        # #     .loc[10:, "Content Type Name"]
        # #     .str.strip()
        # #     .apply(lambda x: " ".join(str(x).split()))
        # # )
        # TC_content_item_list = []
        # degree_sheet = [x for x in content_item_dict.keys() if re.search("^Degree|AreaOfStudy", x, flags=re.I)][0]
        # AreaOfStudy_content_item_list = (
        #     content_item_dict[degree_sheet].iloc[1:, 2].str.strip().apply(lambda x: " ".join(x.split()))
        # )
        # membership_content_item_list = (
        #     content_item_dict["ContentItem-Membership"]
        #     .loc[10:, "Content Type Name"]
        #     .str.strip()
        #     .apply(lambda x: " ".join(str(x).split()))
        # )
        # awards_content_item_list = (
        #     content_item_dict["ContentItem-Honor & Awards"]
        #     .loc[10:, "Content Type Name"]
        #     .str.strip()
        #     .apply(lambda x: " ".join(str(x).split()))
        # )
        # license_content_item_list = (
        #     content_item_dict["ContentItem-License & Certif"]
        #     .loc[10:, "Content Type Name"]
        #     .str.strip()
        #     .apply(lambda x: " ".join(str(x).split()))
        # )

        # WS_df = pd.read_excel(WS_path, sheet_name="Final for SPUR")
        # WS_df.loc[:, "Combined Job"] = (
        #     WS_df.loc[:, "Combined Job"].str.upper().str.strip()
        # )
        # WS_profile_code_list = WS_df.loc[:, "Combined Job"].tolist()
        log_dir_list = os.listdir(log_dir)

        ## Talent Profile ##
        def talent_profile(wb, spur_df, experience_df, job_dat_dir):
            talent_profile = wb.sheets[0]
            talent_profile.range((14, "B"), (10000, "K")).clear()

            # Description
            ur_name_list = []
            profile_code_list = []
            profile_code_missing_list = []
            # print(spur_df)
            for spur_id in spur_df["UR_CODE"].values.tolist():
                if spur_id in experience_df["SPUR ID"].values.tolist():
                    profile_code = (
                        experience_df[experience_df["SPUR ID"] == spur_id]["profile_code"]
                        .drop_duplicates()
                        .values.tolist()
                    )
                    # print(profile_code)
                    # profile_code = list(set([x for x in profile_code]))
                    ur_name = spur_df[spur_df["UR_CODE"] == spur_id]["UR_NAME"].replace("–", "-").values.tolist() * len(
                        profile_code
                    )
                    profile_code_list.extend(profile_code)
                    ur_name_list.extend(ur_name)
                else:
                    profile_code = (
                        spur_df[spur_df["UR_CODE"] == spur_id]["UR_NAME"]
                        .apply(lambda x: profile_code_map(x))
                        .values.tolist()
                    )
                    profile_code = list(set([spur_id + "_" + x for x in profile_code]))
                    ur_name = spur_df[spur_df["UR_CODE"] == spur_id]["UR_NAME"].values.tolist() * len(profile_code)
                    profile_code_list.extend(profile_code)
                    ur_name_list.extend(ur_name)
            profile_code_list = [[x] for x in profile_code_list]
            print(profile_code_list)
            ur_name_list = [[x] for x in ur_name_list]
            # print(profile_code_list)
            # print(ur_name_list)
            talent_profile.range("H14").value = ur_name_list
            row_end = talent_profile.range("H" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            # Column BCEFGJ
            for k in "BCEFGJ":
                talent_profile.range("{}14:{}{}".format(k, k, row_end)).value = "={}$13".format(k)

            # Summary
            talent_profile.range("I14:I{}".format(row_end)).value = "=H14"

            # profile_code = [i + '_' + profile_code_map(j) for i,j in zip(spur_df['UR_CODE'],spur_df['UR_NAME'])]
            talent_profile.range("D14").value = profile_code_list

            # SourceSystemId
            talent_profile.range(
                "K14:K{}".format(row_end)
            ).value = '=CONCAT(UPPER(SUBSTITUTE(D14," ","_")), "_JOB_PROFILE")'

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(talent_profile.range("B1:K1").value).T
            data_range = talent_profile.range("B14:K{}".format(row_end)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            talent_profile_df = pd.DataFrame(data_range)
            # for profile_code in talent_profile_df[2]:
            #     if profile_code not in WS_profile_code_list:
            #         profile_code_missing_list.append(profile_code)
            # if len(profile_code_missing_list) != 0:
            #     logging.warning("[Data validation] {} job code missing".format(len(profile_code_missing_list)))
            #     with open(log_dir + "\\" + f"{process_datetime}_job_code_missing.txt", "w") as f:
            #         f.write("\n".join("{}) ".format(i) + j for i, j in enumerate(profile_code_missing_list, start=1)))
            # else:
            #     if f"{process_datetime}_profile_code_missing.txt" in log_dir_list:
            #         os.remove(log_dir + "\\" + f"{process_datetime}_profile_code_missing.txt")
            talent_profile_df = talent_profile_df[~talent_profile_df[2].isin(profile_code_missing_list)]
            # spur_name_map = dict(
            #     zip(
            #         WS_df.loc[:, "Combined Job"],
            #         WS_df.loc[:, "Finalized Unique Role Title (JobName)"].str.strip(),
            #     )
            # )
            # talent_profile_df.iloc[:, 6] = talent_profile_df.iloc[:, 2].map(spur_name_map)
            # talent_profile_df.iloc[:, 7] = talent_profile_df.iloc[:, 6]

            talent_profile.range((14, "B"), (10000, "K")).clear()
            talent_profile.range((14, "B")).value = talent_profile_df.values.tolist()

            talent_profile_df = pd.concat([header, talent_profile_df])
            talent_profile_df.to_csv(
                job_dat_dir + "\\" + f"TalentProfile.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

        ## Profile Relation ##
        def profile_relation(wb, job_dat_dir):
            profile_relation = wb.sheets[1]
            profile_relation.range((14, "B"), (10000, "K")).clear()
            row_end = wb.sheets[0].range("H" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            # ProfileCode
            profile_relation.range("F14:F{}".format(row_end)).value = "=TalentProfile!D14"

            # Column BCDEGJ
            for k in "BCDEGJ":
                profile_relation.range("{}14:{}{}".format(k, k, row_end)).value = "={}$13".format(k)

            # Column H
            profile_relation.range("H14:H{}".format(row_end)).value = "=F14"

            # Column I
            profile_relation.range("I14:I{}".format(row_end)).value = "PET_SPUR_SET"

            # SourceSystemId
            profile_relation.range(
                "K14:K{}".format(row_end)
            ).value = '=CONCAT(UPPER(SUBSTITUTE(F14," ","_")), "_POSITION_PROFILE_RELATION")'

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(profile_relation.range("B1:K1").value).T
            data_range = profile_relation.range("B14:K{}".format(row_end)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            profile_relation_df = pd.DataFrame(data_range)

            profile_relation_df[2] = "2021/11/01"
            profile_relation_df[3] = "4712/12/31"

            profile_relation.range((14, "B"), (10000, "K")).clear()
            profile_relation.range((14, "B")).value = profile_relation_df.values.tolist()

            profile_relation.range((14, "D"), (row_end, "D")).value = "'2021/11/01"
            profile_relation.range((14, "E"), (row_end, "E")).value = "'4712/12/31"

            profile_relation_df = pd.concat([header, profile_relation_df])
            profile_relation_df.to_csv(
                job_dat_dir + "\\" + f"ProfileRelation.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

        def model_profile_info(wb, job_dat_dir):
            ModelProfileExtraInfo = wb.sheets[2]
            ModelProfileExtraInfo.range((14, "B"), (10000, "I")).clear()
            row_end = wb.sheets[0].range("H" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            # ProfileCode
            ModelProfileExtraInfo.range("D14:D{}".format(row_end)).value = "=TalentProfile!D14"

            # Column BCH
            for k in "BCH":
                ModelProfileExtraInfo.range("{}14:{}{}".format(k, k, row_end)).value = "={}$13".format(k)

            # Column E
            ModelProfileExtraInfo.range("E14:E{}".format(row_end)).value = [
                [re.sub("_\S+", "_DESCRIPTION.txt", x)]
                for x in ModelProfileExtraInfo.range("D14:D{}".format(row_end)).value
            ]

            # Column F
            ModelProfileExtraInfo.range("F14:F{}".format(row_end)).value = [
                [re.sub("_\S+", "_QUALIFICATION.txt", x)]
                for x in ModelProfileExtraInfo.range("D14:D{}".format(row_end)).value
            ]

            # Column G
            ModelProfileExtraInfo.range("G14:G{}".format(row_end)).value = [
                [re.sub("_\S+", "_RESPONSIBILITY.txt", x)]
                for x in ModelProfileExtraInfo.range("D14:D{}".format(row_end)).value
            ]

            # SourceSystemId
            ModelProfileExtraInfo.range(
                "I14:I{}".format(row_end)
            ).value = '=CONCAT(UPPER(SUBSTITUTE(D14," ","_")), "_JOB_PROFILE_MPEI")'

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ModelProfileExtraInfo.range("B1:I1").value).T
            data_range = ModelProfileExtraInfo.range("B14:I{}".format(row_end)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ModelProfileExtraInfo_df = pd.DataFrame(data_range)

            ModelProfileExtraInfo.range((14, "B"), (10000, "I")).clear()
            ModelProfileExtraInfo.range((14, "B")).value = ModelProfileExtraInfo_df.values.tolist()

            ModelProfileExtraInfo_df = pd.concat([header, ModelProfileExtraInfo_df])
            ModelProfileExtraInfo_df.to_csv(
                job_dat_dir + "\\" + f"ModelProfileExtraInfo.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

        def profile_attachment(wb, job_dat_dir):
            ProfileAttachment = wb.sheets[3]
            ProfileAttachment.range((14, "B"), (10000, "K")).clear()
            row_end = wb.sheets[0].range("H" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            # ProfileCode
            ProfileAttachment.range("D14:D{}".format(row_end)).value = "=TalentProfile!H14"

            # Column BCGJ
            for k in "BCGJ":
                ProfileAttachment.range("{}14:{}{}".format(k, k, row_end)).value = "={}$13".format(k)

            # Column F
            ProfileAttachment.range("F14:F{}".format(row_end)).value = "=D14"

            # Column E
            ProfileAttachment.range("E14:E{}".format(row_end)).value = [
                [re.sub("_[A-Za-z]+$", ".pdf", x)] for x in wb.sheets[0].range("D14:D{}".format(row_end)).value
            ]

            # Column H
            ProfileAttachment.range("H14:H{}".format(row_end)).value = "=TalentProfile!D14"

            # Column I
            ProfileAttachment.range("I14:I{}".format(row_end)).value = "=E14"

            # SourceSystemId
            ProfileAttachment.range(
                "K14:K{}".format(row_end)
            ).value = '=CONCAT(UPPER(SUBSTITUTE(H14," ","_")), "_ATTACHMENT")'

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ProfileAttachment.range("B1:K1").value).T
            data_range = ProfileAttachment.range("B14:K{}".format(row_end)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ProfileAttachment_df = pd.DataFrame(data_range)

            ProfileAttachment.range((14, "B"), (10000, "K")).clear()
            ProfileAttachment.range((14, "B")).value = ProfileAttachment_df.values.tolist()

            ProfileAttachment_df = pd.concat([header, ProfileAttachment_df])
            ProfileAttachment_df.to_csv(
                job_dat_dir + "\\" + f"ProfileAttachment.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

        def profile_item_other_descriptor(wb, spur_df, job_dat_dir):
            ProfileItem_OtherDescriptor = wb.sheets[6]
            ProfileItem_OtherDescriptor.range((12, "B"), (10000, "R")).clear()
            row_end = wb.sheets[1].range("H" + str(wb.sheets[1].cells.last_cell.row)).end("up").row

            # ProfileId(SourceSystemId)
            ProfileItem_OtherDescriptor.range("H12:H{}".format(row_end + 1)).value = "=TalentProfile!K11"

            # ProfileId(SourceSystemId)
            ProfileItem_OtherDescriptor.range("J12:J{}".format(row_end + 1)).value = 5

            # Column BCDFGIKLQ
            for k in "BCDFGIKLQ":
                ProfileItem_OtherDescriptor.range("{}12:{}{}".format(k, k, row_end + 1)).value = "={}$11".format(k)

            # SourceSystemId
            ProfileItem_OtherDescriptor.range(
                "R12:R{}".format(row_end + 1)
            ).value = '=CONCAT(UPPER(SUBSTITUTE(TalentProfile!D11," ","_")),"_",L11,"_",UPPER(K12),"_PI")'

            # Challenges
            # print(ProfileItem_OtherDescriptor.range('H12:H{}'.format(row_end+1)).value)
            ProfileItem_OtherDescriptor.range("M12:M{}".format(row_end + 1)).value = [
                [
                    spur_df[spur_df["UR_CODE"] == re.sub("(?<=\d)_.+$", "", x).replace("_", " ")]["CHALLENGES"]
                    .values[0]
                    .replace("\n", "")
                ]
                for x in ProfileItem_OtherDescriptor.range("H12").expand("down").value
            ]

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ProfileItem_OtherDescriptor.range("B2:R2").value).T
            data_range = ProfileItem_OtherDescriptor.range("B12:R{}".format(row_end)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ProfileItem_OtherDescriptor_df = pd.DataFrame(data_range)
            ProfileItem_OtherDescriptor_df[4] = "2021/11/01"
            ProfileItem_OtherDescriptor_df[5] = "4712/12/31"
            ProfileItem_OtherDescriptor_df[8] = ProfileItem_OtherDescriptor_df[8].astype(int)
            ProfileItem_OtherDescriptor_df[9] = ProfileItem_OtherDescriptor_df[9].astype(int)

            ProfileItem_OtherDescriptor.range((12, "B"), (10000, "R")).clear()
            ProfileItem_OtherDescriptor.range((12, "B")).value = ProfileItem_OtherDescriptor_df.values.tolist()

            ProfileItem_OtherDescriptor.range((12, "F"), (row_end, "F")).value = "'2021/11/01"
            ProfileItem_OtherDescriptor.range((12, "G"), (row_end, "G")).value = "'4712/12/31"

            ProfileItem_OtherDescriptor_df = pd.concat([header, ProfileItem_OtherDescriptor_df])
            ProfileItem_OtherDescriptor_df.to_csv(
                job_dat_dir + "\\" + f"ProfileItem-OtherDescriptor.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

        def profile_item_risk(wb, spur_df, job_dat_dir):
            ProfileItem_Risk = wb.sheets[30]
            ProfileItem_Risk.range((12, "B"), (10000, "L")).clear()
            row_end_talent_profile = wb.sheets[0].range("H" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            ## Column G
            ProfileItem_Risk.range((12, "G")).value = [
                ["={}".format(wb.sheets[0].range((row, "K")).get_address(False, False, True))]
                for row in range(14, row_end_talent_profile + 1)
            ]
            row_end_ProfileItem_Risk = (
                ProfileItem_Risk.range("G" + str(ProfileItem_Risk.cells.last_cell.row)).end("up").row
            )

            # Column BCDEFHIK
            for col in "BCDEFHIK":
                ProfileItem_Risk.range((12, col), (row_end_ProfileItem_Risk, col)).value = "={}$11".format(col)

            # SourceSystemId
            ProfileItem_Risk.range(
                "L12:L{}".format(row_end_ProfileItem_Risk + 1)
            ).value = '=CONCATENATE(UPPER(TalentProfile!D14),"_",UPPER(SUBSTITUTE(H12," ","_")),"_PI")'

            # Challenges
            # print(ProfileItem_OtherDescriptor.range('H12:H{}'.format(row_end+1)).value)
            ProfileItem_Risk.range("J12:J{}".format(row_end_ProfileItem_Risk + 1)).value = [
                [
                    spur_df[spur_df["UR_CODE"] == re.sub("(?<=\d)_.+$", "", x).replace("_", " ")]["Challenge"]
                    .values[0]
                    .replace("\n", "")
                ]
                for x in ProfileItem_Risk.range("G12").expand("down").value
            ]

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ProfileItem_Risk.range("B2:L2").value).T
            data_range = ProfileItem_Risk.range("B12:L{}".format(row_end_ProfileItem_Risk)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ProfileItem_Risk_df = pd.DataFrame(data_range)
            ProfileItem_Risk_df[3] = "2021/11/01"
            ProfileItem_Risk_df[4] = "4712/12/31"

            ProfileItem_Risk.range((12, "B"), (10000, "L")).clear()
            ProfileItem_Risk.range((12, "B")).value = ProfileItem_Risk_df.values.tolist()

            ProfileItem_Risk.range((12, "E"), (row_end_ProfileItem_Risk, "E")).value = "'2021/11/01"
            ProfileItem_Risk.range((12, "F"), (row_end_ProfileItem_Risk, "F")).value = "'4712/12/31"

            ProfileItem_Risk_df = pd.concat([header, ProfileItem_Risk_df])
            ProfileItem_Risk_df.to_csv(
                job_dat_dir + "\\" + f"ProfileItem-Risk.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

        def profile_item_exp_required(wb, experience_df, job_dat_dir):
            ProfileItem_ExperienceRequired = wb.sheets[15]
            ProfileItem_ExperienceRequired.range((12, "B"), (10000, "Q")).clear()
            spur_id_occ = experience_df["profile_code"].value_counts().sort_index()

            id_list = []
            exp_importance_list = []
            min_exp_list = []
            max_exp_list = []
            industry_list = []
            domain_list = []
            jg_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(14, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                source_system_id = re.sub("_JOB_PROFILE", "", wb.sheets[0].range((row, "K")).value)
                # source_system_id = re.sub("CHIEF", "Chief", source_system_id)
                # print(source_system_id)

                if source_system_id in spur_id_occ.index:
                    id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(False, False, True)] * spur_id_occ[source_system_id]
                    )

                else:
                    if run_all == True:
                        id_list.extend([wb.sheets[0].range((row, "K")).get_address(False, False, True)])
                    else:
                        continue

                if source_system_id in experience_df["profile_code"].values.tolist():
                    exp_importance_list.extend(
                        experience_df[experience_df["profile_code"] == source_system_id][["Importance"]].values.tolist()
                    )
                    min_exp_list.extend(
                        experience_df[experience_df["profile_code"] == source_system_id][
                            [min_years_column]
                        ].values.tolist()
                    )
                    max_exp_list.extend(
                        experience_df[experience_df["profile_code"] == source_system_id][
                            [max_years_column]
                        ].values.tolist()
                    )
                    
                    industry_list.extend(

                        experience_df[experience_df["profile_code"] == source_system_id][[industry_column]].values.tolist()
                    )
                    if domain_exist == True:
                        domain_list.extend(
                            experience_df[experience_df["profile_code"] == source_system_id][
                                [domain_column]
                            ].values.tolist()
                        )
                    jg_list.extend(
                        experience_df[experience_df["profile_code"] == source_system_id][[exp_jg_column]].values.tolist()
                    )

                else:
                    if run_all == True:
                        exp_importance_list.extend([[""]])
                        min_exp_list.extend([[""]])
                        max_exp_list.extend([[""]])
                        industry_list.extend([[""]])
                        if domain_exist == True:
                            domain_list.extend([[""]])
                        jg_list.extend([[""]])
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
            ProfileItem_ExperienceRequired.range((12, "J")).value = jg_list

            # Column L
            ProfileItem_ExperienceRequired.range((12, "L")).value = min_exp_list

            # Column M
            ProfileItem_ExperienceRequired.range((12, "M")).value = max_exp_list

            # Column N
            if domain_exist == True:
                ProfileItem_ExperienceRequired.range((12, "N")).value = domain_list
            # else:
            #     ProfileItem_ExperienceRequired.range((12, "N")).value = industry_list

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

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ProfileItem_ExperienceRequired.range("B2:Q2").value).T
            data_range = ProfileItem_ExperienceRequired.range("B12:Q{}".format(row_end_ProfileItem_ExperienceRequired)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ProfileItem_ExperienceRequired_df = pd.DataFrame(data_range)
            # duplicated_index = pd.Series()
            ProfileItem_ExperienceRequired_df = ProfileItem_ExperienceRequired_df.drop_duplicates(
                subset=[15], keep="first"
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

            for row in range(12, row_end_ProfileItem_ExperienceRequired_2 + 1):
                if ProfileItem_ExperienceRequired.range((row, "N")).value == None:
                    ProfileItem_ExperienceRequired.range((row, "N")).color = (255, 0, 0)
                if ProfileItem_ExperienceRequired.range((row, "O")).value == None:
                    ProfileItem_ExperienceRequired.range((row, "O")).color = (255, 0, 0)
                # if ProfileItem_ExperienceRequired.range((row, "J")).value == None:
                #     ProfileItem_ExperienceRequired.range((row, "J")).color = (255, 0, 0)
                if ProfileItem_ExperienceRequired.range((row, "L")).value == None:
                    ProfileItem_ExperienceRequired.range((row, "L")).color = (255, 0, 0)
                if ProfileItem_ExperienceRequired.range((row, "M")).value == None:
                    ProfileItem_ExperienceRequired.range((row, "M")).color = (255, 0, 0)

            ProfileItem_ExperienceRequired_df = pd.concat([header, ProfileItem_ExperienceRequired_df])
            ProfileItem_ExperienceRequired_df.to_csv(
                job_dat_dir + "\\" + f"ProfileItem-ExperienceRequired.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
                line_terminator="\n",
            )

        def profile_item_competency_LC(wb, leadership_competency_df, job_dat_dir):
            ProfileItem_Competency_LC = wb.sheets[19]
            ProfileItem_Competency_LC.range((13, "B"), (100000, "Q")).clear()
            spur_id_leadership_competency_occ = leadership_competency_df["SPUR ID"].value_counts().sort_index()
            lc_id_list = []
            lc_list = []
            min_list = []
            max_list = []
            
            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(14, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
                #     print(ur_id)

                if ur_id in spur_id_leadership_competency_occ.index:
                    lc_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_leadership_competency_occ[ur_id]
                    )

                else:
                    if run_all == True:
                        lc_id_list.extend([wb.sheets[0].range((row, "K")).get_address(True, False, True)])
                    else:
                        continue

                if ur_id in leadership_competency_df["SPUR ID"].values.tolist():
                    lc_list.extend(leadership_competency_df[leadership_competency_df["SPUR ID"] == ur_id][[leadership_competency_column]].values.tolist())
                    min_list.extend(leadership_competency_df[leadership_competency_df["SPUR ID"] == ur_id][['MinimumProficiency']].values.tolist())
                    max_list.extend(leadership_competency_df[leadership_competency_df["SPUR ID"] == ur_id][['MaximumProficiency']].values.tolist())
                else:
                    continue

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
            ProfileItem_Competency_LC_df[4] = "2021/11/01"
            ProfileItem_Competency_LC_df[5] = "4712/12/31"

            ProfileItem_Competency_LC_df[[9, 11]] = ProfileItem_Competency_LC_df[[9, 11]].apply(
                pd.to_numeric, downcast="signed"
            )

            ProfileItem_Competency_LC.range((13, "B"), (100000, "O")).clear()
            ProfileItem_Competency_LC.range((13, "B")).value = ProfileItem_Competency_LC_df.values.tolist()

            # Column F
            ProfileItem_Competency_LC.range((13, "F"), (last_row_lc, "F")).value = "'2021/11/01"

            # Column G
            ProfileItem_Competency_LC.range((13, "G"), (last_row_lc, "G")).value = "'4712/12/31"

            ProfileItem_Competency_LC_df = pd.concat([header, ProfileItem_Competency_LC_df])
            ProfileItem_Competency_LC_df.to_csv(
                job_dat_dir + "\\" + f"ProfileItem-Competency_LC.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

            # LC_content_item_list = LC_content_item_list.values.tolist()
            # for row in range(13, last_row_lc + 1):
            #     if not ProfileItem_Competency_LC.range((row, "E")).value in LC_content_item_list:
            #         ProfileItem_Competency_LC.range((row, "E")).color = (255, 0, 0)

        def profile_item_competency_TC(wb, technical_competency_df, job_dat_dir):
            ProfileItem_Competency_TC = wb.sheets[20]
            ProfileItem_Competency_TC.range((13, "B"), (10000, "P")).clear()

            spur_id_technical_competency_occ = technical_competency_df["SPUR ID"].value_counts().sort_index()
            tc_id_list = []
            tc_list = []
            min_list = []
            max_list = []
            importance_list = []
            
            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(14, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
                #     print(ur_id)

                if ur_id in spur_id_technical_competency_occ.index:
                    tc_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_technical_competency_occ[ur_id]
                    )

                else:
                    if run_all == True:
                        tc_id_list.extend([wb.sheets[0].range((row, "K")).get_address(True, False, True)])
                    else:
                        continue

                if ur_id in technical_competency_df["SPUR ID"].values.tolist():
                    tc_list.extend(technical_competency_df[technical_competency_df["SPUR ID"] == ur_id][[technical_competency_column]].values.tolist())
                    min_list.extend(technical_competency_df[technical_competency_df["SPUR ID"] == ur_id][['MinimumProficiency']].values.tolist())
                    max_list.extend(technical_competency_df[technical_competency_df["SPUR ID"] == ur_id][['MaximumProficiency']].values.tolist())
                    importance_list.extend(technical_competency_df[technical_competency_df["SPUR ID"] == ur_id][['Importance']].values.tolist())
                else:
                    continue
            # print(min_tc_list)
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

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ProfileItem_Competency_TC.range("B2:P2").value).T
            data_range = ProfileItem_Competency_TC.range("B13:P{}".format(last_row_tc)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ProfileItem_Competency_TC_df = pd.DataFrame(data_range).dropna(subset=[10])
            ProfileItem_Competency_TC_df = ProfileItem_Competency_TC_df[
                (ProfileItem_Competency_TC_df[3].notna()) | (ProfileItem_Competency_TC_df[3] == "")
            ]
            ProfileItem_Competency_TC_df[4] = "2021/11/01"
            ProfileItem_Competency_TC_df[5] = "4712/12/31"

            ProfileItem_Competency_TC.range((13, "B"), (10000, "P")).clear()
            ProfileItem_Competency_TC.range((13, "B")).value = ProfileItem_Competency_TC_df.values.tolist()
            last_row_tc_2 = (
                ProfileItem_Competency_TC.range("H" + str(ProfileItem_Competency_TC.cells.last_cell.row)).end("up").row
            )

            # Column F
            ProfileItem_Competency_TC.range((13, "F"), (last_row_tc_2, "F")).value = "'2021/11/01"

            # Column G
            ProfileItem_Competency_TC.range((13, "G"), (last_row_tc_2, "G")).value = "'4712/12/31"

            ProfileItem_Competency_TC_df[[10, 12]] = ProfileItem_Competency_TC_df[[10, 12]].apply(
                pd.to_numeric, downcast="signed"
            )
            ProfileItem_Competency_TC_df[8] = ProfileItem_Competency_TC_df[8].fillna(-1).astype(int).replace(-1, "")
            ProfileItem_Competency_TC_df[10] = ProfileItem_Competency_TC_df[10].fillna(-1).astype(int).replace(-1, "")
            ProfileItem_Competency_TC_df[12] = ProfileItem_Competency_TC_df[12].fillna(-1).astype(int).replace(-1, "")

            ProfileItem_Competency_TC_df = pd.concat([header, ProfileItem_Competency_TC_df])
            ProfileItem_Competency_TC_df.to_csv(
                job_dat_dir + "\\" + f"ProfileItem-Competency_TC.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

            # TC_content_item_list = TC_content_item_list.values.tolist()
            # for row in range(13, last_row_tc_2 + 1):
            #     if not ProfileItem_Competency_TC.range((row, "E")).value in TC_content_item_list:
            #         ProfileItem_Competency_TC.range((row, "E")).color = (255, 0, 0)
            #     if ProfileItem_Competency_TC.range((row, "L")).value == None:
            #         ProfileItem_Competency_TC.range((row, "L")).color = (255, 0, 0)
            #         ProfileItem_Competency_TC.range((row, "N")).color = (255, 0, 0)

        def profile_item_degree(wb, degree_df, job_dat_dir):
            ProfileItem_Degree = wb.sheets[7]
            ProfileItem_Degree.range((12, "B"), (10000, "Q")).clear()

            spur_id_degree_occ = degree_df["SPUR ID"].value_counts().sort_index()

            degree_id_list = []
            degree_importance_list = []
            edu_level_list = []
            degree_name_list = []
            country_code_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(14, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
                #     print(ur_id)

                if ur_id in spur_id_degree_occ.index:
                    degree_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_degree_occ[ur_id]
                    )

                else:
                    if run_all == True:
                        degree_id_list.extend([wb.sheets[0].range((row, "K")).get_address(True, False, True)])
                    else:
                        continue

                if ur_id in degree_df["SPUR ID"].values.tolist():
                    degree_importance_list.extend(degree_df[degree_df["SPUR ID"] == ur_id][["Importance"]].values.tolist())
                    edu_level_list.extend(degree_df[degree_df["SPUR ID"] == ur_id][[degree_column]].values.tolist())
                    degree_name_list.extend(
                        degree_df[degree_df["SPUR ID"] == ur_id][[area_of_study_column]].values.tolist()
                    )
                    country_code_list.extend(
                        degree_df[degree_df["SPUR ID"] == ur_id][['CountryCode']].values.tolist()
                    )
                else:
                    if run_all == True:
                        degree_importance_list.extend([[""]])
                        edu_level_list.extend([[""]])
                        degree_name_list.extend([[""]])
                        country_code_list.extend([[""]])
                    else:
                        continue

            # Column H
            ProfileItem_Degree.range((12, "H")).value = [["={}".format(k)] for k in degree_id_list]

            # Column K
            # ProfileItem_Degree.range((12, "K")).value = degree_importance_list

            # Column E
            ProfileItem_Degree.range((12, "E")).value = edu_level_list

            # Column O
            ProfileItem_Degree.range((12, "O")).value = degree_name_list

            # Column M
            ProfileItem_Degree.range((12, "M")).value = country_code_list

            # Column N
            ProfileItem_Degree.range((12, "N")).value = country_code_list

            row_end_degree_sheet = (
                ProfileItem_Degree.range("H" + str(ProfileItem_Degree.cells.last_cell.row)).end("up").row
            )
            # Column BCDFGILMNP
            for k in "BCDFGILP":
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
            ProfileItem_Degree.range((12, "Q")).value = formula_list

            # Drop competency with minimum proficiency = 0
            header = pd.DataFrame(ProfileItem_Degree.range("B2:Q2").value).T
            data_range = ProfileItem_Degree.range("B12:Q{}".format(row_end_degree_sheet)).value
            if isinstance(data_range, list) and not isinstance(data_range[0], list):
                data_range = [data_range]
            ProfileItem_Degree_df = pd.DataFrame(data_range)
            ProfileItem_Degree_df[4] = "2021/11/01"
            ProfileItem_Degree_df[5] = "4712/12/31"
            # ProfileItem_Degree_df[9] = ProfileItem_Degree_df[9].astype(int)

            ProfileItem_Degree.range((12, "B"), (10000, "Q")).clear()
            ProfileItem_Degree.range((12, "B")).value = ProfileItem_Degree_df.values.tolist()

            ProfileItem_Degree.range((12, "F"), (row_end_degree_sheet, "F")).value = "'2021/11/01"
            ProfileItem_Degree.range((12, "G"), (row_end_degree_sheet, "G")).value = "'4712/12/31"

            # AreaOfStudy_content_item_list = AreaOfStudy_content_item_list.values.tolist()
            # AreaOfStudy_content_item_list = [x.strip() for x in AreaOfStudy_content_item_list]
            # for row in range(12, row_end_degree_sheet + 1):
            #     if not ProfileItem_Degree.range((row, "O")).value in AreaOfStudy_content_item_list:
            #         ProfileItem_Degree.range((row, "O")).color = (255, 0, 0)
                # if ProfileItem_Degree.range((row, "K")).value == None:
                #     ProfileItem_Degree.range((row, "K")).color = (255, 0, 0)
                #     ProfileItem_Degree.range((row, "Q")).color = (255, 0, 0)

            ProfileItem_Degree_df = pd.concat([header, ProfileItem_Degree_df])
            ProfileItem_Degree_df.to_csv(
                job_dat_dir + "\\" + f"ProfileItem-Degree.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
                float_format="%.f",
            )

        def profile_item_language(wb, language_df, job_dat_dir):
            ProfileItem_Language = wb.sheets[11]
            ProfileItem_Language.range((12, "B"), (10000, "R")).clear()

            spur_id_language_occ = language_df["SPUR ID"].value_counts().sort_index()
            language_id_list = []
            language_list = []
            reading_proficiency_list = []
            writing_proficiency_list = []
            speaking_proficiency_list = []
            required_list = []
            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row

            for row in range(14, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)
                #     print(ur_id)

                if ur_id in spur_id_language_occ.index:
                    language_id_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_language_occ[ur_id]
                    )

                else:
                    if run_all == True:
                        language_id_list.extend([wb.sheets[0].range((row, "K")).get_address(True, False, True)])
                    else:
                        continue

                if ur_id in language_df["SPUR ID"].values.tolist():
                    language_list.extend(language_df[language_df["SPUR ID"] == ur_id][[language_column]].values.tolist())
                    reading_proficiency_list.extend(language_df[language_df["SPUR ID"] == ur_id][['ReadingProficiency']].values.tolist())
                    writing_proficiency_list.extend(language_df[language_df["SPUR ID"] == ur_id][['WritingProficiency']].values.tolist())
                    speaking_proficiency_list.extend(language_df[language_df["SPUR ID"] == ur_id][['SpeakingProficiency']].values.tolist())
                    required_list.extend(language_df[language_df["SPUR ID"] == ur_id][['Required']].values.tolist())

                else:
                    continue

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
            ProfileItem_Language_df[4] = "2021/11/01"
            ProfileItem_Language_df[5] = "4712/12/31"
            # ProfileItem_Language_df[15] = ProfileItem_Language_df[15].astype(int)

            ProfileItem_Language.range((12, "B"), (10000, "R")).clear()
            ProfileItem_Language.range((12, "B")).value = ProfileItem_Language_df.values.tolist()

            ProfileItem_Language.range((12, "F"), (row_end_language_sheet, "F")).value = "'2021/11/01"
            ProfileItem_Language.range((12, "G"), (row_end_language_sheet, "G")).value = "'4712/12/31"

            ProfileItem_Language_df = pd.concat([header, ProfileItem_Language_df])
            print(ProfileItem_Language_df)
            ProfileItem_Language_df.to_csv(
                job_dat_dir + "\\" + f"ProfileItem-Language.dat",
                header=None,
                index=None,
                sep="|",
                mode="w",
            )

        def profile_item_membership(wb, membership_df, job_dat_dir):
            ProfileItem_Membership = wb.sheets[13]
            ProfileItem_Membership.range((12, "B"), (10000, "M")).clear()

            spur_id_membership_occ = membership_df["SPUR ID"].value_counts().sort_index()

            talent_profile_list = []
            membership_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(14, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)

                if ur_id in spur_id_membership_occ.index:
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_membership_occ[ur_id]
                    )

                else:
                    continue

                if ur_id in membership_df["SPUR ID"].values.tolist():
                    membership_list.extend(
                        membership_df[membership_df["SPUR ID"] == ur_id][[membership_column]].values.tolist()
                    )

                else:
                    continue

            if membership_list != []:
                # Column H
                ProfileItem_Membership.range((12, "H")).value = [["={}".format(k)] for k in talent_profile_list]

                # Column E
                ProfileItem_Membership.range((12, "E")).value = membership_list

                row_end_membership_sheet = (
                    ProfileItem_Membership.range("H" + str(wb.sheets[13].cells.last_cell.row)).end("up").row
                )
                # Column BCDFGIJKL
                for k in "BCDFGIJKL":
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
                ProfileItem_Membership.range((12, "M")).value = formula_list

                # membership_content_item_list = membership_content_item_list.values.tolist()
                # for row in range(12, row_end_membership_sheet + 1):
                #     if not ProfileItem_Membership.range((row, "E")).value in membership_content_item_list:
                #         ProfileItem_Membership.range((row, "E")).color = (255, 0, 0)

                # Drop competency with minimum proficiency = 0
                header = pd.DataFrame(ProfileItem_Membership.range("B2:M2").value).T
                data_range = ProfileItem_Membership.range("B12:M{}".format(row_end_membership_sheet)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_Membership_df = pd.DataFrame(data_range)
                ProfileItem_Membership_df[4] = "2021/11/01"
                ProfileItem_Membership_df[5] = "4712/12/31"
                # ProfileItem_Membership_df[9] = ProfileItem_Membership_df[9].astype(int)
                ProfileItem_Membership_df = ProfileItem_Membership_df.drop_duplicates(subset=[11], keep="first")

                ProfileItem_Membership.range((12, "B"), (10000, "M")).clear()
                ProfileItem_Membership.range((12, "B")).value = ProfileItem_Membership_df.values.tolist()
                row_end_membership_sheet_2 = (
                    ProfileItem_Membership.range("H" + str(wb.sheets[13].cells.last_cell.row)).end("up").row
                )

                ProfileItem_Membership.range((12, "F"), (row_end_membership_sheet_2, "F")).value = "'2021/11/01"
                ProfileItem_Membership.range((12, "G"), (row_end_membership_sheet_2, "G")).value = "'4712/12/31"

                ProfileItem_Membership_df = pd.concat([header, ProfileItem_Membership_df])
                ProfileItem_Membership_df.to_csv(
                    job_dat_dir + "\\" + f"ProfileItem-Membership.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                )

            else:
                pass

        def profile_item_awards(wb, awards_df, job_dat_dir):
            ProfileItem_Awards = wb.sheets[9]
            ProfileItem_Awards.range((12, "B"), (10000, "N")).clear()

            spur_id_awards_occ = awards_df["SPUR ID"].value_counts().sort_index()

            talent_profile_list = []
            awards_list = []
            importance_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(14, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)

                if ur_id in spur_id_awards_occ.index:
                    talent_profile_list.extend(
                        [wb.sheets[1].range((row, "K")).get_address(True, False, True)] * spur_id_awards_occ[ur_id]
                    )

                else:
                    continue

                if ur_id in awards_df["SPUR ID"].values.tolist():
                    awards_list.extend(awards_df[awards_df["SPUR ID"] == ur_id][["Honor & Awards"]].values.tolist())
                    importance_list.extend(awards_df[awards_df["SPUR ID"] == ur_id][["Importance"]].values.tolist())

                else:
                    continue

            if awards_list != []:
                # Column H
                ProfileItem_Awards.range((12, "H")).value = [["={}".format(k)] for k in talent_profile_list]

                # Column E
                ProfileItem_Awards.range((12, "E")).value = awards_list

                # Column J
                # ProfileItem_Awards.range((12, "J")).value = importance_list

                row_end_awards_sheet = ProfileItem_Awards.range("H" + str(wb.sheets[9].cells.last_cell.row)).end("up").row
                # Column BCDFGIJKM
                for k in "BCDFGIJKLM":
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
                ProfileItem_Awards.range((12, "N")).value = formula_list

                # awards_content_item_list = awards_content_item_list.values.tolist()
                # for row in range(12, row_end_awards_sheet + 1):
                #     if not ProfileItem_Awards.range((row, "E")).value in awards_content_item_list:
                #         ProfileItem_Awards.range((row, "E")).color = (255, 0, 0)

                # Drop competency with minimum proficiency = 0
                header = pd.DataFrame(ProfileItem_Awards.range("B2:N2").value).T
                data_range = ProfileItem_Awards.range("B12:N{}".format(row_end_awards_sheet)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_Awards_df = pd.DataFrame(data_range)
                ProfileItem_Awards_df[4] = "2021/11/01"
                ProfileItem_Awards_df[5] = "4712/12/31"
                # ProfileItem_Awards_df[8] = ProfileItem_Awards_df[8].astype(int)
                ProfileItem_Awards_df = ProfileItem_Awards_df.drop_duplicates(subset=[12], keep="first")

                ProfileItem_Awards.range((12, "B"), (10000, "N")).clear()
                ProfileItem_Awards.range((12, "B")).value = ProfileItem_Awards_df.values.tolist()
                row_end_awards_sheet_2 = (
                    ProfileItem_Awards.range("H" + str(wb.sheets[21].cells.last_cell.row)).end("up").row
                )

                ProfileItem_Awards.range((12, "F"), (row_end_awards_sheet_2, "F")).value = "'2021/11/01"
                ProfileItem_Awards.range((12, "G"), (row_end_awards_sheet_2, "G")).value = "'4712/12/31"

                ProfileItem_Awards_df = pd.concat([header, ProfileItem_Awards_df])
                ProfileItem_Awards_df.to_csv(
                    job_dat_dir + "\\" + f"ProfileItem-Awards.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                )

            else:
                pass

        def profile_item_license(wb, license_df, job_dat_dir):
            ProfileItem_License = wb.sheets[5]
            ProfileItem_License.range((12, "B"), (10000, "Q")).clear()

            spur_id_license_occ = license_df["SPUR ID"].value_counts().sort_index()

            talent_profile_list = []
            license_list = []
            importance_list = []
            required_list = []
            country_code_list = []
            state_name_list = []

            row_end_talent_profile = wb.sheets[0].range("D" + str(wb.sheets[0].cells.last_cell.row)).end("up").row
            for row in range(14, row_end_talent_profile + 1):
                if wb.sheets[0].range((row, "K")).value == None:
                    break

                ur_id = re.sub("_.+", "", wb.sheets[0].range((row, "K")).value)

                if ur_id in spur_id_license_occ.index:
                    talent_profile_list.extend(
                        [wb.sheets[0].range((row, "K")).get_address(True, False, True)] * spur_id_license_occ[ur_id]
                    )

                else:
                    continue

                if ur_id in license_df["SPUR ID"].values.tolist():
                    license_list.extend(license_df[license_df["SPUR ID"] == ur_id][[license_column]].values.tolist())
                    importance_list.extend(license_df[license_df["SPUR ID"] == ur_id][["Importance"]].values.tolist())
                    required_list.extend(license_df[license_df["SPUR ID"] == ur_id][['Required']].values.tolist())
                    country_code_list.extend(license_df[license_df["SPUR ID"] == ur_id][['Country']].values.tolist())
                    state_name_list.extend(license_df[license_df["SPUR ID"] == ur_id][['State']].values.tolist())
                else:
                    continue

            if license_list != []:
                # Column H
                ProfileItem_License.range((12, "H")).value = [["={}".format(k)] for k in talent_profile_list]

                # Column E
                ProfileItem_License.range((12, "E")).value = license_list

                # Column K
                ProfileItem_License.range((12, "K")).value = ""

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

                # license_content_item_list = license_content_item_list.values.tolist()

                # Drop competency with minimum proficiency = 0
                header = pd.DataFrame(ProfileItem_License.range("B2:Q2").value).T
                data_range = ProfileItem_License.range("B12:Q{}".format(row_end_license_sheet)).value
                if isinstance(data_range, list) and not isinstance(data_range[0], list):
                    data_range = [data_range]
                ProfileItem_License_df = pd.DataFrame(data_range)
                ProfileItem_License_df[4] = "2021/11/01"
                ProfileItem_License_df[5] = "4712/12/31"
                # ProfileItem_License_df[8] = ProfileItem_License_df[8].astype(int)
                ProfileItem_License_df = ProfileItem_License_df.drop_duplicates(subset=[15], keep="first")

                ProfileItem_License.range((12, "B"), (10000, "Q")).clear()
                ProfileItem_License.range((12, "B")).value = ProfileItem_License_df.values.tolist()
                row_end_license_sheet_2 = (
                    ProfileItem_License.range("H" + str(wb.sheets[5].cells.last_cell.row)).end("up").row
                )

                ProfileItem_License.range((12, "F"), (row_end_license_sheet_2, "F")).value = "'2021/11/01"
                ProfileItem_License.range((12, "G"), (row_end_license_sheet_2, "G")).value = "'4712/12/31"

                # for row in range(12, row_end_license_sheet_2 + 1):
                #     if not ProfileItem_License.range((row, "E")).value in license_content_item_list:
                #         ProfileItem_License.range((row, "E")).color = (255, 0, 0)

                ProfileItem_License_df = pd.concat([header, ProfileItem_License_df])
                ProfileItem_License_df.to_csv(
                    job_dat_dir + "\\" + f"ProfileItem-License.dat",
                    header=None,
                    index=None,
                    sep="|",
                    mode="w",
                )

            else:
                pass

        # Execute functions
        log.info("[Job profile] Talent Profile")
        talent_profile(wb, spur_df, experience_df, job_dat_dir)

        log.info("[Job profile] Profile Relation")
        profile_relation(wb, job_dat_dir)

        log.info("[Job profile] Model Profile Info")
        model_profile_info(wb, job_dat_dir)

        log.info("[Job profile] Profile Attachment")
        profile_attachment(wb, job_dat_dir)

        # log.info("[Job profile] Profile Item Other Descriptor")
        # profile_item_other_descriptor(wb, spur_df, job_dat_dir)

        log.info("[Job profile] License & Certificate")
        profile_item_license(wb, license_df, job_dat_dir)

        log.info("[Job profile] Degree")
        profile_item_degree(wb, degree_df, job_dat_dir)

        log.info("[Job profile] Honors & Awards")
        profile_item_awards(wb, awards_df, job_dat_dir)

        log.info("[Job profile] Language")
        profile_item_language(wb, language_df, job_dat_dir)

        log.info("[Job profile] Membership")
        profile_item_membership(wb, membership_df, job_dat_dir)

        log.info("[Job profile] Experience Required")
        profile_item_exp_required(wb, experience_df, job_dat_dir)

        log.info("[Job profile] Leadership Competency")
        profile_item_competency_LC(wb, leadership_competency_df, job_dat_dir)

        log.info("[Job profile] Technical Competency")
        profile_item_competency_TC(wb, technical_competency_df, job_dat_dir)

        log.info("[Job profile] Profile Item Risk")
        profile_item_risk(wb, spur_df, job_dat_dir)

    except Exception as e:
        raise ValueError(e)
    finally:
        if 'wb' in locals():
            wb.save()
            wb.close()