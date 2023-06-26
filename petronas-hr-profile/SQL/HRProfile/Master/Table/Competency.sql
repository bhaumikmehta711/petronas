CREATE TABLE [Master].[Competency]
(
	[CompetencyID] INT IDENTITY(1, 1) NOT NULL,
	[CompetencyName] VARCHAR(100) NOT NULL,
	[CompetencyType] VARCHAR(20) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_Competency_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME CONSTRAINT [DC_Competency_ModifiedTimeStamp] DEFAULT(GETUTCDATE()) NOT NULL,
	CONSTRAINT [PK_Competency_CompetencyID] PRIMARY KEY ([CompetencyID])
)
