CREATE TABLE [Master].[Award]
(
	[AwardID] INT IDENTITY(1, 1) NOT NULL,
	[AwardName] VARCHAR(500) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_Award_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_Award_AwardID] PRIMARY KEY ([AwardID]),
	CONSTRAINT [UC_Award_AwardName] UNIQUE ([AwardName])
)