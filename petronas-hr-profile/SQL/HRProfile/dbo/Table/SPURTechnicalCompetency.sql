CREATE TABLE [dbo].[SPURTechnicalCompetency]
(
	[SPURID] INT NOT NULL,
	[TechnicalCompetencyID] INT NOT NULL,
	[MinimumProficiencyLevel] VARCHAR(10),
	[MaxiumProficiencyLevel] VARCHAR(10),
	[Importance] VARCHAR(10),
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_SPURTechnicalCompetency_SPURID_TechnicalCompetencyID] PRIMARY KEY ([SPURID], [TechnicalCompetencyID]),
	CONSTRAINT [FK_SPUR_SPURTechnicalCompetency_SPURID] FOREIGN KEY ([SPURID]) REFERENCES [DBO].[SPUR]([SPURID]),
)
