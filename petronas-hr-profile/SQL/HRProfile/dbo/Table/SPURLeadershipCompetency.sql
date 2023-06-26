CREATE TABLE [dbo].[SPURLeadershipCompetency]
(
	[SPURID] INT NOT NULL,
	[LeadershipCompetencyID] INT NOT NULL,
	[MinimumProficiencyLevel] VARCHAR(10),
	[MaxiumProficiencyLevel] VARCHAR(10),
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_SPURLeadershipCompetency_SPURID_LeadershipCompetencyID] PRIMARY KEY ([SPURID], [LeadershipCompetencyID]),
	CONSTRAINT [FK_SPUR_SPURLeadershipCompetency_SPURID] FOREIGN KEY ([SPURID]) REFERENCES [DBO].[SPUR]([SPURID]),
)
