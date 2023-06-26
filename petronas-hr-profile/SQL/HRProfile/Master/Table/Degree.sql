CREATE TABLE [Master].[Degree]
(
	[DegreeID] INT IDENTITY(1, 1) NOT NULL,
	[DegreeName] VARCHAR(500) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_Degree_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_Degree_LanguageID] PRIMARY KEY ([DegreeID]),
	CONSTRAINT [UC_Degree_LanguageName] UNIQUE ([DegreeName])
)