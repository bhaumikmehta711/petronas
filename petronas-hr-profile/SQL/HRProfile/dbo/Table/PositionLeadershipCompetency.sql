CREATE TABLE [dbo].[PositionLeadershipCompetency]
(
	[PositionID] INT NOT NULL,
	[LeadershipCompetencyID] INT NOT NULL,
	[MinimumProficiencyLevel] VARCHAR(10),
	[MaxiumProficiencyLevel] VARCHAR(10),
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_PositionLeadershipCompetency_PositionID_LeadershipCompetencyID] PRIMARY KEY ([PositionID], [LeadershipCompetencyID]),
	CONSTRAINT [FK_Position_PositionLeadershipCompetency_PositionID] FOREIGN KEY ([PositionID]) REFERENCES [DBO].[Position]([PositionID]),
)
