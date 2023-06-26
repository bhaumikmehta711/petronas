CREATE TABLE [dbo].[PositionTechnicalCompetency]
(
	[PositionID] INT NOT NULL,
	[TechnicalCompetencyID] INT NOT NULL,
	[MinimumProficiencyLevel] VARCHAR(10),
	[MaxiumProficiencyLevel] VARCHAR(10),
	[Importance] VARCHAR(10),
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_PositionTechnicalCompetency_PositionID_TechnicalCompetencyID] PRIMARY KEY ([PositionID], [TechnicalCompetencyID]),
	CONSTRAINT [FK_Position_PositionTechnicalCompetency_PositionID] FOREIGN KEY ([PositionID]) REFERENCES [DBO].[Position]([PositionID]),
)
