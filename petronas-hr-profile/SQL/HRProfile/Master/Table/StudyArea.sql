CREATE TABLE [Master].[StudyArea]
(
	[StudyAreaID] INT IDENTITY(1, 1) NOT NULL,
	[StudyAreaName] VARCHAR(500) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_StudyArea_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_StudyArea_StudyAreaID] PRIMARY KEY ([StudyAreaID]),
	CONSTRAINT [UC_StudyArea_StudyAreaName] UNIQUE ([StudyAreaName])
)