from pandas import ExcelWriter
import logging.config
from Helper.sql_helper import *

log = logging.getLogger(__name__)


def trim_all_columns(df):
    """
    Trim whitespace from ends of each value across all series in dataframe
    """
    trim_strings = lambda x: x.strip() if isinstance(x, str) else x
    return df.applymap(trim_strings)


class data_processor:
    """
    docstring
    """

    def __init__(
        self,
        consumption_dir,
        process_datetime,
        sql_engine
    ):
        self.consumption_dir = consumption_dir
        self.process_datetime = process_datetime
        self.sql_engine = sql_engine


    def position_data(self):
        license_df = sql_read(self.sql_engine, f"SELECT \
                E.PositionCode [PositionID],\
                B.LicenseName [License & Certi (e.g. Transportation Management Certificate)],\
                '' [if Other (Specific)], \
                '' [JG],\
                '' [Importance],\
                C.CountryCode [Country],\
                D.StateName [State],\
                CASE WHEN A.Required = 1 THEN \'Y\' ELSE \'N\' END [Required] \
            FROM [PositionLicense] A \
            INNER JOIN [Master].[License] B ON A.LicenseID = B.LicenseID\
            INNER JOIN [Master].[Country] C ON A.CountryID = C.CountryID\
            INNER JOIN [Master].[State] D ON A.StateID = D.StateID\
            INNER JOIN [dbo].[Position] E ON E.PositionID = A.PositionID\
            INNER JOIN [Staging].[Position_{self.process_datetime}] F ON F.PositionID = E.PositionID\
        ")

        exp_df = sql_read(self.sql_engine, f"SELECT \
                B.PositionCode [PositionID],\
                CAST(A.MinimumExperienceRequired AS VARCHAR(10)) [mimimumExperienceRequired],\
                CAST(A.DesiredExperienceRequired AS VARCHAR(10)) [Desired Years Of experience],\
                A.Industry [Industry],\
                A.Domain [Domain],\
                C.RoleLevelCode [Role Level],\
                '' [JG],\
                '' [Importance]\
            FROM [PositionExperience] A \
            INNER JOIN [Position] B ON A.PositionID = B.PositionID \
            INNER JOIN [Master].[RoleLevel] C ON C.RoleLevelID = B.RoleLevelID\
            INNER JOIN [Staging].[Position_{self.process_datetime}] D ON D.PositionID = B.PositionID\
        ")

        degree_df = sql_read(self.sql_engine, f"SELECT \
                E.PositionCode [PositionID],\
                B.DegreeName [ContentItem],\
                C.StudyAreaName [AreaOfStudy], \
                D.CountryCode [CountryCode], \
                CASE WHEN A.Required = 1 THEN \'Y\' ELSE \'N\' END [Required], \
                '' [if Other (Specific)], \
                '' [JG],\
                '' [Importance]\
            FROM [PositionDegree] A \
            INNER JOIN [Master].[Degree] B ON A.DegreeID = B.DegreeID\
            INNER JOIN [Master].[StudyArea] C ON C.StudyAreaID = A.StudyAreaID\
            INNER JOIN [Master].[Country] D ON D.CountryID = A.CountryID\
            INNER JOIN [dbo].[Position] E ON E.PositionID = A.PositionID\
            INNER JOIN [Staging].[Position_{self.process_datetime}] F ON F.PositionID = E.PositionID\
        ")

        membership_df = sql_read(self.sql_engine, f"SELECT \
                C.PositionCode [PositionID],\
                B.MembershipName [Bodies membership Name (e.g. Board of Engineering Malaysia)],\
                CASE WHEN A.Required = 1 THEN \'Y\' ELSE \'N\' END [Required], \
                Title [Title], \
                '' [if Other (Specific)], \
                '' [JG],\
                '' [Importance]\
            FROM [PositionMembership] A \
            INNER JOIN [Master].[Membership] B ON A.MembershipID = B.MembershipID\
            INNER JOIN [dbo].[Position] C ON C.PositionID = A.PositionID\
            INNER JOIN [Staging].[Position_{self.process_datetime}] D ON D.PositionID = C.PositionID\
        ")

        language_df = sql_read(self.sql_engine, f"SELECT \
                F.PositionCode [PositionID],\
                B.LanguageName [Language],\
                C.LanguageProficiencyCode ReadingProficiency, \
                D.LanguageProficiencyCode WritingProficiency, \
                E.LanguageProficiencyCode SpeakingProficiency, \
                CASE WHEN A.Required = 1 THEN \'Y\' ELSE \'N\' END Required \
            FROM [PositionLanguage] A \
            INNER JOIN [Master].[Language] B ON A.LanguageID = B.LanguageID\
            INNER JOIN [Master].[LanguageProficiency] C ON C.LanguageProficiencyID = A.ReadingLanguageProficiencyID\
            INNER JOIN [Master].[LanguageProficiency] D ON D.LanguageProficiencyID = A.WritingLanguageProficiencyID\
            INNER JOIN [Master].[LanguageProficiency] E ON E.LanguageProficiencyID = A.SpeakingLanguageProficiencyID\
            INNER JOIN [dbo].[Position] F ON F.PositionID = A.PositionID\
            INNER JOIN [Staging].[Position_{self.process_datetime}] G ON G.PositionID = F.PositionID\
        ")

        awards_df = sql_read(self.sql_engine, f"SELECT \
                C.PositionCode [PositionID],\
                B.AwardName [Honor & Awards],\
                '' [if Other (Specific)], \
                CASE WHEN A.Required = 1 THEN \'Y\' ELSE \'N\' END [Required], \
                '' [JG],\
                '' [Importance]\
            FROM [PositionAward] A \
            INNER JOIN [Master].[Award] B ON A.AwardID = B.AwardID\
            INNER JOIN [dbo].[Position] C ON C.PositionID = A.PositionID\
            INNER JOIN [Staging].[Position_{self.process_datetime}] D ON D.PositionID = C.PositionID\
        ")

        leadership_competency_df = sql_read(self.sql_engine, f"SELECT \
                E.PositionCode [PositionID],\
                B.LeadershipCompetencyName [Competency],\
                CAST(ISNULL(C.LeadershipCompetencyProficiencyValue, \'\') AS VARCHAR(10)) MaximumProficiency,\
                CAST(ISNULL(D.LeadershipCompetencyProficiencyValue, \'\') AS VARCHAR(10)) MinimumProficiency\
            FROM [PositionLeadershipCompetency] A \
            INNER JOIN [Master].[LeadershipCompetency] B ON B.LeadershipCompetencyID = A.LeadershipCompetencyID\
            INNER JOIN [Master].[LeadershipCompetencyProficiency] C ON C.LeadershipCompetencyProficiencyID = A.MaximumLeadershipCompetencyProficiencyID\
            INNER JOIN [Master].[LeadershipCompetencyProficiency] D ON D.LeadershipCompetencyProficiencyID = A.MinimumLeadershipCompetencyProficiencyID\
            INNER JOIN [Position] E ON E.PositionID = A.PositionID\
            INNER JOIN [Staging].[Position_{self.process_datetime}] F ON F.PositionID = E.PositionID\
        ")

        technical_competency_df = sql_read(self.sql_engine, f"SELECT \
                F.PositionCode [PositionID],\
                B.TechnicalCompetencyName [Competency],\
                CAST(ISNULL(C.TechnicalCompetencyProficiencyValue, \'\') AS VARCHAR(10)) MaximumProficiency,\
                CAST(ISNULL(D.TechnicalCompetencyProficiencyValue, \'\') AS VARCHAR(10)) MinimumProficiency,\
				CAST(ISNULL(E.CompetencyImportanceValue, \'\') AS VARCHAR(10)) Importance\
            FROM [PositionTechnicalCompetency] A \
            INNER JOIN [Master].[TechnicalCompetency] B ON B.TechnicalCompetencyID = A.TechnicalCompetencyID\
            INNER JOIN [Master].[TechnicalCompetencyProficiency] C ON C.TechnicalCompetencyProficiencyID = A.MaximumTechnicalCompetencyProficiencyID\
            INNER JOIN [Master].[TechnicalCompetencyProficiency] D ON D.TechnicalCompetencyProficiencyID = A.MinimumTechnicalCompetencyProficiencyID\
			INNER JOIN [Master].[CompetencyImportance] E ON E.CompetencyImportanceID = A.PositionTechnicalCompetencyID\
            INNER JOIN [Position] F ON F.PositionID = A.PositionID\
            INNER JOIN [Staging].[Position_{self.process_datetime}] G ON G.PositionID = F.PositionID")

        a = [
            (exp_df, "Experience"),
            (degree_df, "Degree"),
            (membership_df, "Membership"),
            (awards_df, "Awards"),
            (license_df, "License"),
            (language_df, "Language"),
            (leadership_competency_df, "LeadershipCompetency"),
            (technical_competency_df, "TechnicalCompetency")
        ]
        with ExcelWriter(
            self.consumption_dir + "\\" + "data\\final_processed_data\\{}_position_profile_data.xlsx".format(self.process_datetime),
            mode="w",
            engine="openpyxl",
        ) as writer:
            for df in a:
                df[0].to_excel(writer, sheet_name=df[1], index=False)