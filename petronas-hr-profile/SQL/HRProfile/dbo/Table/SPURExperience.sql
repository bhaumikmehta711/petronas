CREATE TABLE [dbo].[SPURExperience]
(
	[SPURExperienceID] INT IDENTITY(1, 1) NOT NULL,
	[SPURID] INT NOT NULL,
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
	CONSTRAINT [PK_SPURExperience_SPURID_ExperienceID] PRIMARY KEY ([SPURExperienceID]),
	CONSTRAINT [FK_SPUR_SPURExperience_SPURID] FOREIGN KEY ([SPURID]) REFERENCES [DBO].[SPUR]([SPURID])
)