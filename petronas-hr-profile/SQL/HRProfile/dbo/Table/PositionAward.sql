CREATE TABLE [dbo].[PositionAward]
(
	[PositionID] INT NOT NULL,
	[AwardID] INT NOT NULL,
	[Establishment] VARCHAR(500),
	[Required] BIT,
	[Importance] VARCHAR(50) NULL,
	[JG] VARCHAR(100) NULL,
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_PositionAward_PositionID_AwardID] PRIMARY KEY ([PositionID], [AwardID]),
	CONSTRAINT [FK_Position_PositionAward_PositionID] FOREIGN KEY ([PositionID]) REFERENCES [DBO].[Position]([PositionID]),
	CONSTRAINT [FK_Language_PositionAward_AwardID] FOREIGN KEY ([AwardID]) REFERENCES [Master].[Award]([AwardID])
)
