CREATE TABLE [Master].[LeadershipCompetency]
(
	[LeadershipCompetencyID] INT IDENTITY(1, 1) NOT NULL,
	[LeadershipCompetencyName] VARCHAR(500) NOT NULL,
	[RoleName] VARCHAR(500) NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_LeadershipCompetency_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_LeadershipCompetency_LeadershipCompetencyID] PRIMARY KEY ([LeadershipCompetencyID]),
	CONSTRAINT [UC_LeadershipCompetency_LeadershipCompetencyName] UNIQUE ([LeadershipCompetencyName])
)