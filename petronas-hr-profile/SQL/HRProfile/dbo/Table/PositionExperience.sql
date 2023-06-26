CREATE TABLE [dbo].[PositionExperience]
(
	[PositionExperienceID] INT IDENTITY(1, 1) NOT NULL,
	[PositionID] INT NOT NULL,
	[Skill] VARCHAR(500),
	[MinimumExperienceRequired] DECIMAL(4,2),
	[MaximumExperienceRequired] DECIMAL(4,2),
	[RecommendedTenure] DECIMAL(4,2),
	[Industry] VARCHAR(500),
	[Domain] VARCHAR(500),
	[Importance] VARCHAR(50) NULL,
	[JG] VARCHAR(50) NULL,
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_PositionExperience_PositionID_ExperienceID] PRIMARY KEY ([PositionExperienceID]),
	CONSTRAINT [FK_Position_PositionExperience_PositionID] FOREIGN KEY ([PositionID]) REFERENCES [DBO].[Position]([PositionID])
)