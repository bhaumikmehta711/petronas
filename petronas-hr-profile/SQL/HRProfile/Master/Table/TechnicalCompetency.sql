CREATE TABLE [Master].[TechnicalCompetency]
(
	[TechnicalCompetencyID] INT IDENTITY(1, 1) NOT NULL,
	[TechnicalCompetencyCode] VARCHAR(500) NOT NULL,
	[TechnicalCompetencyName] VARCHAR(500) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_TechnicalCompetency_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_TechnicalCompetency_TechnicalCompetencyID] PRIMARY KEY ([TechnicalCompetencyID]),
	CONSTRAINT [UC_TechnicalCompetency_TechnicalCompetencyName] UNIQUE ([TechnicalCompetencyName])
)