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
        try:
            self.consumption_dir = consumption_dir
            self.process_datetime = process_datetime
            self.sql_engine = sql_engine
        except Exception as e:
            raise ValueError(e)


    def spur_data(self):
        try:
            license_df = sql_read(self.sql_engine, f"SELECT \
                    A.ProfileCode,\
                    C.LicenseName,\
                    ISNULL(D.CountryCode, \'\') CountryCode,\
                    ISNULL(E.StateName, \'\') StateName,\
                    ISNULL(B.Title, \'\') Title,\
                    CASE WHEN B.Required = 1 THEN 'Y' ELSE 'N' END [Required] \
                FROM [Staging].[SPUR_{self.process_datetime}] A \
                INNER JOIN [SPURLicense] B ON B.SPURID = A.SPURID\
                INNER JOIN [Master].[License] C ON C.LicenseID = B.LicenseID\
                LEFT JOIN [Master].[Country] D ON D.CountryID = B.CountryID\
                LEFT JOIN [Master].[State] E ON E.StateID = B.StateID\
            ")

            exp_df = sql_read(self.sql_engine, f"SELECT \
                    A.ProfileCode,\
                    CAST(B.MinimumExperienceRequired AS VARCHAR(10)) [MimimumExperienceRequired],\
                    CAST(B.DesiredExperienceRequired AS VARCHAR(10)) [MaximumExperienceRequired],\
                    ISNULL(B.Industry, \'\') [Industry],\
                    ISNULL(B.Domain, \'\') [Domain],\
                    ISNULL(B.Skill, \'\') Skill\
                FROM [Staging].[SPUR_{self.process_datetime}] A\
                INNER JOIN [SPURExperience] B ON A.SPURID = B.SPURID\
            ")

            degree_df = sql_read(self.sql_engine, f"SELECT \
                    A.ProfileCode,\
                    C.DegreeName,\
                    ISNULL(D.StudyAreaName, \'\') StudyAreaName, \
                    ISNULL(B.Major, \'\') Major,\
                    ISNULL(F.SchoolName, \'\') School,\
                    ISNULL(E.CountryCode, \'\') CountryCode\
                FROM [Staging].[SPUR_{self.process_datetime}] A\
                INNER JOIN [SPURDegree] B ON B.SPURID = A.SPURID\
                INNER JOIN [Master].[Degree] C ON C.DegreeID = B.DegreeID\
                LEFT JOIN [Master].[StudyArea] D ON D.StudyAreaID = B.StudyAreaID\
                LEFT JOIN [Master].[Country] E ON E.CountryID = B.CountryID\
                LEFT JOIN [Master].[School] F ON F.SchoolID = B.SchoolID\
            ")

            membership_df = sql_read(self.sql_engine, f"SELECT \
                    A.ProfileCode,\
                    C.MembershipName,\
                    ISNULL(B.Title, \'\') Title\
                FROM [Staging].[SPUR_{self.process_datetime}] A\
                INNER JOIN [SPURMembership] B ON B.SPURID = A.SPURID\
                INNER JOIN [Master].[Membership] C ON C.MembershipID = B.MembershipID\
            ")

            language_df = sql_read(self.sql_engine, f"SELECT \
                    A.ProfileCode,\
                    C.LanguageName,\
                    ISNULL(D.LanguageProficiencyCode, \'\') ReadingProficiency, \
                    ISNULL(E.LanguageProficiencyCode, \'\') WritingProficiency, \
                    ISNULL(F.LanguageProficiencyCode, \'\') SpeakingProficiency, \
                    CASE WHEN B.Required = 1 THEN 'Y' ELSE 'N' END Required \
                FROM [Staging].[SPUR_{self.process_datetime}] A\
                INNER JOIN [SPURLanguage] B ON B.SPURID = A.SPURID\
                INNER JOIN [Master].[Language] C ON C.LanguageID = B.LanguageID\
                LEFT JOIN [Master].[LanguageProficiency] D ON D.LanguageProficiencyID = B.ReadingLanguageProficiencyID\
                LEFT JOIN [Master].[LanguageProficiency] E ON E.LanguageProficiencyID = B.WritingLanguageProficiencyID\
                LEFT JOIN [Master].[LanguageProficiency] F ON F.LanguageProficiencyID = B.SpeakingLanguageProficiencyID\
            ")

            awards_df = sql_read(self.sql_engine, f"SELECT \
                    A.ProfileCode,\
                    C.AwardName,\
                    CASE WHEN B.Required = 1 THEN 'Y' ELSE 'N' END Required \
                FROM [Staging].[SPUR_{self.process_datetime}] A\
                INNER JOIN [SPURAward] B ON B.SPURID = A.SPURID\
                INNER JOIN [Master].[Award] C ON C.AwardID = B.AwardID\
            ")

            leadership_competency_df = sql_read(self.sql_engine, f"SELECT \
                    A.ProfileCode,\
                    ISNULL(C.LeadershipCompetencyName, \'\') LeadershipCompetencyName,\
                    CAST(ISNULL(D.LeadershipCompetencyProficiencyValue, '0') AS VARCHAR(10)) MaximumProficiency,\
                    CAST(ISNULL(E.LeadershipCompetencyProficiencyValue, '0') AS VARCHAR(10)) MinimumProficiency\
                FROM [Staging].[SPUR_{self.process_datetime}] A\
                INNER JOIN [SPURLeadershipCompetency] B ON B.SPURID = A.SPURID \
                INNER JOIN [Master].[LeadershipCompetency] C ON C.LeadershipCompetencyID = B.LeadershipCompetencyID\
                LEFT JOIN [Master].[LeadershipCompetencyProficiency] D ON D.LeadershipCompetencyProficiencyID = B.MaximumLeadershipCompetencyProficiencyID\
                LEFT JOIN [Master].[LeadershipCompetencyProficiency] E ON E.LeadershipCompetencyProficiencyID = B.MinimumLeadershipCompetencyProficiencyID\
            ")

            technical_competency_df = sql_read(self.sql_engine, f"SELECT \
                    A.ProfileCode,\
                    C.TechnicalCompetencyName,\
                    CAST(ISNULL(D.TechnicalCompetencyProficiencyValue, '0') AS VARCHAR(10)) MaximumProficiency,\
                    CAST(ISNULL(E.TechnicalCompetencyProficiencyValue, '0') AS VARCHAR(10)) MinimumProficiency,\
                    CAST(ISNULL(F.CompetencyImportanceValue, '') AS VARCHAR(10)) Importance\
                FROM [Staging].[SPUR_{self.process_datetime}] A\
                INNER JOIN [SPURTechnicalCompetency] B ON B.SPURID = A.SPURID\
                INNER JOIN [Master].[TechnicalCompetency] C ON C.TechnicalCompetencyID = B.TechnicalCompetencyID\
                LEFT JOIN [Master].[TechnicalCompetencyProficiency] D ON D.TechnicalCompetencyProficiencyID = B.MaximumTechnicalCompetencyProficiencyID\
                LEFT JOIN [Master].[TechnicalCompetencyProficiency] E ON E.TechnicalCompetencyProficiencyID = B.MinimumTechnicalCompetencyProficiencyID\
                LEFT JOIN [Master].[CompetencyImportance] F ON F.CompetencyImportanceID = B.TechnicalCompetencyImportanceID"\
            )

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
                self.consumption_dir + "\\" + "data\\final_processed_data\\{}_details.xlsx".format(self.process_datetime),
                mode="w",
                engine="openpyxl",
            ) as writer:
                for df in a:
                    df[0].to_excel(writer, sheet_name=df[1], index=False)
        except Exception as e:
            raise ValueError(e)