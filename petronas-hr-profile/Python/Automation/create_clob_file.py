def create_clob_file(
    df,
    clob_folder_path
):
    for i in range(len(df)):
        # DESCRIPTION
        with open(
            clob_folder_path + "\\" + df.iloc[i]['ProfileCode'] + "_DESCRIPTION.txt",
            "w",
            encoding="utf-8",
        ) as f:
            f.write(df.iloc[i]["PurposeAndAccountability"].strip())
        
        # RESPONSIBILITY
        with open(
            clob_folder_path + "\\" + df.iloc[i]['ProfileCode'] + "_RESPONSIBILITY.txt",
            "w",
            encoding="utf-8",
        ) as f:
            f.write(df.iloc[i]["KPI"].strip())
        
        # QUALIFICATION
        with open(
            clob_folder_path + "\\" + df.iloc[i]['ProfileCode'] + "_QUALIFICATION.txt",
            "w",
            encoding="utf-8",
        ) as f:
            f.write(df.iloc[i]["Experience"].strip())