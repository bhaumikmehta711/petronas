import pandas as pd
from pandas import ExcelWriter
import numpy as np
import re
import itertools
import glob
import os
import sys
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
import openpyxl
import string
from openpyxl import Workbook, load_workbook
import logging.config
from utility import sql_read

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


    def spur_data(self):
        license_df = sql_read(self.sql_engine, f"SELECT \
                C.SPURCode [SPUR ID],\
                B.LicenseName [License & Certi (e.g. Transportation Management Certificate)],\
                '' [if Other (Specific)], \
                A.[JG],\
                A.[Importance]\
            FROM [SPURLicense] A \
            INNER JOIN [Master].[License] B ON A.LicenseID = B.LicenseID\
            INNER JOIN [dbo].[SPUR] C ON C.SPURID = A.SPURID\
        ")

        exp_df = sql_read(self.sql_engine, f"SELECT \
                B.SPURCode [SPUR ID],\
                A.MinimumExperienceRequired [mimimumExperienceRequired],\
                A.DesiredExperienceRequired [Desired Years Of experience],\
                A.Industry [Industry],\
                A.Domain [Domain],\
                A.JG [JG],\
                A.Importance [Importance]\
            FROM [SPURExperience] A \
            INNER JOIN [SPUR] B ON A.SPURID = B.SPURID \
        ")

        degree_df = sql_read(self.sql_engine, f"SELECT \
                D.SPURCode [SPUR ID],\
                B.DegreeName [ContentItem],\
                C.StudyAreaName [AreaOfStudy], \
                '' [if Other (Specific)], \
                A.[JG],\
                A.[Importance]\
            FROM [SPURDegree] A \
            INNER JOIN [Master].[Degree] B ON A.DegreeID = B.DegreeID\
            INNER JOIN [Master].[StudyArea] C ON C.StudyAreaID = A.StudyAreaID\
            INNER JOIN [dbo].[SPUR] D ON D.SPURID = A.SPURID\
        ")

        membership_df = sql_read(self.sql_engine, f"SELECT \
                C.SPURCode [SPUR ID],\
                B.MembershipName [Bodies membership Name (e.g. Board of Engineering Malaysia)],\
                '' [if Other (Specific)], \
                A.[JG],\
                A.[Importance]\
            FROM [SPURMembership] A \
            INNER JOIN [Master].[Membership] B ON A.MembershipID = B.MembershipID\
            INNER JOIN [dbo].[SPUR] C ON C.SPURID = A.SPURID\
        ")

        awards_df = sql_read(self.sql_engine, f"SELECT \
                C.SPURCode [SPUR ID],\
                B.AwardName [Honor & Awards Name (e.g. Long Service Award - 10 years)],\
                '' [if Other (Specific)], \
                A.[JG],\
                A.[Importance]\
            FROM [SPURAward] A \
            INNER JOIN [Master].[Award] B ON A.AwardID = B.AwardID\
            INNER JOIN [dbo].[SPUR] C ON C.SPURID = A.SPURID\
        ")

        a = [
            (exp_df, "Experience"),
            (degree_df, "Degree"),
            (membership_df, "Membership"),
            (awards_df, "Awards"),
            (license_df, "License"),
        ]
        with ExcelWriter(
            self.consumption_dir + "\\" + "data\\final_processed_data\\{}_details.xlsx".format(self.process_datetime),
            mode="w",
            engine="openpyxl",
        ) as writer:
            for df in a:
                df[0].to_excel(writer, sheet_name=df[1], index=False)